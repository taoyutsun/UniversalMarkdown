import ctypes
import json
import os
import subprocess
import sys
from importlib.metadata import entry_points
from typing import Optional

from ctypes import POINTER, Structure, byref, c_void_p, cast
from ctypes import wintypes

os.environ["QT_LOGGING_RULES"] = "qt.qpa.window=false"

from PyQt6.QtCore import QThread, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from converter import DocxToMarkdownConverter

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

try:
    from markitdown import MarkItDown
    MARKITDOWN_IMPORT_ERROR = None
except Exception as exc:
    MarkItDown = None
    MARKITDOWN_IMPORT_ERROR = exc


APP_TITLE = "Universal Markdown 極簡轉檔器"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def get_runtime_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return BASE_DIR


RUNTIME_BASE_DIR = get_runtime_base_dir()
DOTENV_PATH = os.path.join(RUNTIME_BASE_DIR, ".env")

DEFAULT_PROVIDER = "NVIDIA_NIM"
DEFAULT_DOCX_MODE = "mammoth"
DEFAULT_OCR_MODE = "fast"

DOCX_MODE_OPTIONS = [
    ("Mammoth 保留圖片", "mammoth"),
    ("MarkItDown OCR", "markitdown"),
]

OCR_MODE_OPTIONS = [
    ("快速模式", "fast"),
    ("OCR 增強模式", "ocr"),
]

MAMMOTH_EXTENSIONS = {".docx"}
MARKITDOWN_EXTENSIONS = {
    ".pdf",
    ".pptx",
    ".xlsx",
    ".xls",
    ".csv",
    ".txt",
    ".json",
    ".xml",
    ".html",
    ".htm",
    ".epub",
}

PROVIDER_STARTER_TEMPLATES = [
    {
        "name": "NVIDIA_NIM",
        "base_url": "https://integrate.api.nvidia.com/v1",
        "model": "meta/llama-3.2-90b-vision-instruct",
    },
    {
        "name": "Google_Gemini",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "model": "gemini-3-flash-preview",
    },
]

SETTINGS_DIR = os.path.join(RUNTIME_BASE_DIR, "_UserSettings")
SETTINGS_PATH = os.path.join(SETTINGS_DIR, "settings.json")
LEGACY_SETTINGS_PATHS = []
if os.getenv("APPDATA"):
    LEGACY_SETTINGS_PATHS.append(
        os.path.join(os.getenv("APPDATA"), "UniversalMarkdownConverter", "settings.json")
    )
ACTIVE_SETTINGS_PATH = SETTINGS_PATH

USER_SETTINGS: dict = {}

CRED_TYPE_GENERIC = 1
CRED_PERSIST_LOCAL_MACHINE = 2
ERROR_NOT_FOUND = 1168
CREDENTIAL_COMMENT = "Universal Markdown Converter API Key"


LPBYTE = POINTER(ctypes.c_ubyte)


class FILETIME(Structure):
    _fields_ = [
        ("dwLowDateTime", wintypes.DWORD),
        ("dwHighDateTime", wintypes.DWORD),
    ]


class CREDENTIAL_ATTRIBUTEW(Structure):
    _fields_ = [
        ("Keyword", wintypes.LPWSTR),
        ("Flags", wintypes.DWORD),
        ("ValueSize", wintypes.DWORD),
        ("Value", LPBYTE),
    ]


PCREDENTIAL_ATTRIBUTEW = POINTER(CREDENTIAL_ATTRIBUTEW)


class CREDENTIALW(Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD),
        ("Type", wintypes.DWORD),
        ("TargetName", wintypes.LPWSTR),
        ("Comment", wintypes.LPWSTR),
        ("LastWritten", FILETIME),
        ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", LPBYTE),
        ("Persist", wintypes.DWORD),
        ("AttributeCount", wintypes.DWORD),
        ("Attributes", PCREDENTIAL_ATTRIBUTEW),
        ("TargetAlias", wintypes.LPWSTR),
        ("UserName", wintypes.LPWSTR),
    ]


PCREDENTIALW = POINTER(CREDENTIALW)


try:
    _advapi32 = ctypes.WinDLL("Advapi32.dll", use_last_error=True) if sys.platform == "win32" else None
except Exception:
    _advapi32 = None


if _advapi32 is not None:
    _advapi32.CredWriteW.argtypes = [PCREDENTIALW, wintypes.DWORD]
    _advapi32.CredWriteW.restype = wintypes.BOOL
    _advapi32.CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, POINTER(PCREDENTIALW)]
    _advapi32.CredReadW.restype = wintypes.BOOL
    _advapi32.CredDeleteW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
    _advapi32.CredDeleteW.restype = wintypes.BOOL
    _advapi32.CredFree.argtypes = [c_void_p]
    _advapi32.CredFree.restype = None


def exception_hook(exctype, value, traceback):
    print("應用程式發生未處理例外：")
    import traceback as tb

    tb.print_exception(exctype, value, traceback)
    sys.exit(1)


sys.excepthook = exception_hook


def build_starter_provider_profiles() -> list[dict]:
    return [
        {
            "name": "NVIDIA_NIM",
            "base_url": os.getenv("NVIDIA_NIM_BASE_URL", "https://integrate.api.nvidia.com/v1").strip(),
            "api_key": "",
            "model": "meta/llama-3.2-90b-vision-instruct",
        },
        {
            "name": "Google_Gemini",
            "base_url": os.getenv(
                "GEMINI_BASE_URL",
                "https://generativelanguage.googleapis.com/v1beta/openai/",
            ).strip(),
            "api_key": "",
            "model": "gemini-3-flash-preview",
        },
    ]


def get_starter_profile_by_name(provider_name: str) -> Optional[dict]:
    provider_name = (provider_name or "").strip()
    for profile in build_starter_provider_profiles():
        if profile.get("name") == provider_name:
            return profile
    return None


def normalize_provider_profile(raw: dict, fallback: Optional[dict] = None) -> dict:
    fallback = fallback or {}
    name = str(raw.get("name", fallback.get("name", ""))).strip()
    base_url = str(raw.get("base_url", fallback.get("base_url", ""))).strip()
    api_key = str(raw.get("api_key", fallback.get("api_key", ""))).strip()
    model = str(raw.get("model", fallback.get("model", ""))).strip()

    if not base_url:
        base_url = str(fallback.get("base_url", "")).strip()
    if not api_key:
        api_key = str(fallback.get("api_key", "")).strip()
    if not model:
        model = str(fallback.get("model", "")).strip()

    return {
        "name": name,
        "base_url": base_url,
        "api_key": api_key,
        "model": model,
    }


def normalize_provider_profiles(raw) -> list[dict]:
    starter_profiles = build_starter_provider_profiles()
    starter_map = {profile["name"]: profile for profile in starter_profiles}
    normalized: list[dict] = []

    if isinstance(raw, list):
        for item in raw:
            if not isinstance(item, dict):
                continue
            fallback = starter_map.get(str(item.get("name", "")).strip(), {})
            profile = normalize_provider_profile(item, fallback=fallback)
            if profile["name"]:
                normalized.append(profile)
    elif isinstance(raw, dict):
        for provider_name, provider_values in raw.items():
            if not isinstance(provider_values, dict):
                continue
            fallback = starter_map.get(provider_name, {"name": provider_name})
            profile = normalize_provider_profile(
                {"name": provider_name, **provider_values},
                fallback=fallback,
            )
            if profile["name"]:
                normalized.append(profile)

    return normalized or starter_profiles


def supports_credential_manager() -> bool:
    return sys.platform == "win32" and _advapi32 is not None


def credential_target_name(provider_name: str) -> str:
    return f"UniversalMarkdownConverter/{provider_name.strip()}"


def read_api_key_from_credential_manager(provider_name: str) -> Optional[str]:
    if not supports_credential_manager() or not provider_name.strip():
        return None

    credential_ptr = PCREDENTIALW()
    target_name = credential_target_name(provider_name)
    success = _advapi32.CredReadW(target_name, CRED_TYPE_GENERIC, 0, byref(credential_ptr))
    if not success:
        error_code = ctypes.get_last_error()
        if error_code == ERROR_NOT_FOUND:
            return None
        raise RuntimeError(f"讀取 Windows Credential Manager 失敗：{ctypes.FormatError(error_code)}")

    try:
        credential = credential_ptr.contents
        if not credential.CredentialBlob or credential.CredentialBlobSize <= 0:
            return ""
        raw = ctypes.string_at(credential.CredentialBlob, credential.CredentialBlobSize)
        return raw.decode("utf-16-le")
    finally:
        _advapi32.CredFree(credential_ptr)


def write_api_key_to_credential_manager(provider_name: str, api_key: str):
    if not supports_credential_manager() or not provider_name.strip():
        return

    target_name = credential_target_name(provider_name)
    credential = CREDENTIALW()
    blob = api_key.encode("utf-16-le")
    blob_buffer = ctypes.create_string_buffer(blob)

    credential.Flags = 0
    credential.Type = CRED_TYPE_GENERIC
    credential.TargetName = target_name
    credential.Comment = CREDENTIAL_COMMENT
    credential.CredentialBlobSize = len(blob)
    credential.CredentialBlob = cast(blob_buffer, LPBYTE) if blob else None
    credential.Persist = CRED_PERSIST_LOCAL_MACHINE
    credential.AttributeCount = 0
    credential.Attributes = None
    credential.TargetAlias = None
    credential.UserName = provider_name

    success = _advapi32.CredWriteW(byref(credential), 0)
    if not success:
        error_code = ctypes.get_last_error()
        raise RuntimeError(f"寫入 Windows Credential Manager 失敗：{ctypes.FormatError(error_code)}")


def delete_api_key_from_credential_manager(provider_name: str):
    if not supports_credential_manager() or not provider_name.strip():
        return

    target_name = credential_target_name(provider_name)
    success = _advapi32.CredDeleteW(target_name, CRED_TYPE_GENERIC, 0)
    if not success:
        error_code = ctypes.get_last_error()
        if error_code == ERROR_NOT_FOUND:
            return
        raise RuntimeError(f"刪除 Windows Credential Manager 失敗：{ctypes.FormatError(error_code)}")


def attach_stored_api_keys(profiles: list[dict]) -> list[dict]:
    attached_profiles: list[dict] = []
    for profile in profiles:
        hydrated = dict(profile)
        stored_api_key = read_api_key_from_credential_manager(hydrated.get("name", ""))
        if stored_api_key:
            hydrated["api_key"] = stored_api_key
        attached_profiles.append(hydrated)
    return attached_profiles


def sync_credentials_from_settings(settings: dict, previous_settings: Optional[dict] = None):
    if not supports_credential_manager():
        return

    previous_names = set(get_provider_names(previous_settings or {}))
    current_names = set(get_provider_names(settings))

    for profile in settings.get("providers", []):
        if not isinstance(profile, dict):
            continue
        provider_name = str(profile.get("name", "")).strip()
        api_key = str(profile.get("api_key", "")).strip()
        if not provider_name:
            continue
        if api_key:
            # 同名 Provider 會覆寫原本那一筆 Credential，不會新增重複項。
            write_api_key_to_credential_manager(provider_name, api_key)
        else:
            delete_api_key_from_credential_manager(provider_name)

    for removed_name in previous_names - current_names:
        delete_api_key_from_credential_manager(removed_name)


def serialize_settings_for_disk(settings: dict) -> dict:
    sanitized_providers = []
    for profile in settings.get("providers", []):
        if not isinstance(profile, dict):
            continue
        profile_copy = normalize_provider_profile(profile)
        if supports_credential_manager():
            profile_copy.pop("api_key", None)
        sanitized_providers.append(profile_copy)

    return {
        "providers": sanitized_providers,
        "ui": dict(settings.get("ui", {})),
        "meta": dict(settings.get("meta", {})),
    }


def get_env_api_key_for_provider(provider_name: str) -> str:
    provider_name = (provider_name or "").strip()
    if provider_name == "NVIDIA_NIM":
        return os.getenv("NVIDIA_NIM_API_KEY", "").strip() or os.getenv("NVIDIA_API_KEY", "").strip()
    if provider_name == "Google_Gemini":
        return os.getenv("GEMINI_API_KEY", "").strip() or os.getenv("GOOGLE_API_KEY", "").strip()
    return ""


def maybe_migrate_env_api_keys(settings: dict) -> bool:
    if not supports_credential_manager():
        return False

    meta = settings.setdefault("meta", {})
    if meta.get("env_api_migrated"):
        return False

    providers = settings.get("providers", [])
    if isinstance(providers, list):
        for profile in providers:
            if not isinstance(profile, dict):
                continue
            provider_name = str(profile.get("name", "")).strip()
            if not provider_name:
                continue
            if read_api_key_from_credential_manager(provider_name):
                continue
            env_api_key = get_env_api_key_for_provider(provider_name)
            if env_api_key:
                write_api_key_to_credential_manager(provider_name, env_api_key)

    meta["env_api_migrated"] = True
    return True


def default_user_settings() -> dict:
    return {
        "providers": build_starter_provider_profiles(),
        "ui": {
            "provider": DEFAULT_PROVIDER,
            "docx_mode": DEFAULT_DOCX_MODE,
            "ocr_mode": DEFAULT_OCR_MODE,
        },
        "meta": {
            "env_api_migrated": False,
        },
    }


def get_settings_display_path() -> str:
    return ACTIVE_SETTINGS_PATH


def select_best_settings_source() -> tuple[Optional[str], Optional[dict]]:
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as settings_file:
                return SETTINGS_PATH, json.load(settings_file)
        except Exception:
            pass

    best_path = None
    best_payload = None
    best_score = -1

    for candidate_path in LEGACY_SETTINGS_PATHS:
        if not os.path.exists(candidate_path):
            continue
        try:
            with open(candidate_path, "r", encoding="utf-8") as settings_file:
                payload = json.load(settings_file)
        except Exception:
            continue

        providers = payload.get("providers", [])
        has_api_key_on_disk = any(
            isinstance(profile, dict) and bool(str(profile.get("api_key", "")).strip())
            for profile in providers
        ) if isinstance(providers, list) else False
        meta = payload.get("meta", {})
        is_migrated = isinstance(meta, dict) and bool(meta.get("env_api_migrated"))

        score = 0
        if is_migrated:
            score += 2
        if not has_api_key_on_disk:
            score += 1

        if score > best_score:
            best_score = score
            best_path = candidate_path
            best_payload = payload

    return best_path, best_payload


def load_user_settings() -> dict:
    global ACTIVE_SETTINGS_PATH

    settings = default_user_settings()
    candidate_path, loaded = select_best_settings_source()

    if loaded is None:
        ACTIVE_SETTINGS_PATH = SETTINGS_PATH
        settings["providers"] = attach_stored_api_keys(settings["providers"])
        return settings

    ACTIVE_SETTINGS_PATH = SETTINGS_PATH

    if isinstance(loaded, dict):
        settings["providers"] = attach_stored_api_keys(
            normalize_provider_profiles(loaded.get("providers"))
        )
        ui_settings = loaded.get("ui", {})
        if isinstance(ui_settings, dict):
            settings["ui"].update(ui_settings)
        meta_settings = loaded.get("meta", {})
        if isinstance(meta_settings, dict):
            settings["meta"].update(meta_settings)

    if candidate_path and candidate_path != SETTINGS_PATH:
        try:
            save_user_settings(settings)
        except Exception:
            ACTIVE_SETTINGS_PATH = candidate_path

    return settings


def save_user_settings(settings: dict):
    global ACTIVE_SETTINGS_PATH
    payload = serialize_settings_for_disk(settings)
    os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
    with open(SETTINGS_PATH, "w", encoding="utf-8") as settings_file:
        json.dump(payload, settings_file, ensure_ascii=False, indent=2)
    ACTIVE_SETTINGS_PATH = SETTINGS_PATH


def get_provider_names(settings: Optional[dict] = None) -> list[str]:
    settings = settings or USER_SETTINGS
    providers = settings.get("providers", [])
    if not isinstance(providers, list):
        return []
    return [
        provider["name"]
        for provider in providers
        if isinstance(provider, dict) and provider.get("name")
    ]


def get_provider_profile(provider_name: str, settings: Optional[dict] = None) -> Optional[dict]:
    settings = settings or USER_SETTINGS
    providers = settings.get("providers", [])
    if not isinstance(providers, list):
        return None
    for provider in providers:
        if isinstance(provider, dict) and provider.get("name") == provider_name:
            return provider
    return None


def get_default_provider_name(settings: Optional[dict] = None) -> str:
    names = get_provider_names(settings)
    if DEFAULT_PROVIDER in names:
        return DEFAULT_PROVIDER
    return names[0] if names else DEFAULT_PROVIDER


def get_user_ui_value(key: str, default: str) -> str:
    ui_settings = USER_SETTINGS.get("ui", {})
    value = ui_settings.get(key, default)
    return value if isinstance(value, str) and value else default


def reload_environment_settings():
    if load_dotenv is not None:
        load_dotenv(DOTENV_PATH, override=True)

    global USER_SETTINGS
    settings = load_user_settings()
    if maybe_migrate_env_api_keys(settings):
        save_user_settings(settings)
        settings = load_user_settings()
    USER_SETTINGS = settings


reload_environment_settings()


def is_placeholder(value: str) -> bool:
    return (not value) or ("example.com" in value) or value.startswith("PASTE_")


def get_file_extension(file_path: str) -> str:
    return os.path.splitext(file_path)[1].lower()


def resolve_engine(file_path: str) -> str:
    extension = get_file_extension(file_path)
    if extension in MAMMOTH_EXTENSIONS:
        return "docx"
    if extension in MARKITDOWN_EXTENSIONS:
        return "markitdown"
    return "unsupported"


def has_ocr_plugin() -> bool:
    try:
        return any(ep.name == "ocr" for ep in entry_points(group="markitdown.plugin"))
    except Exception:
        return False


def build_llm_provider(provider_name: Optional[str] = None) -> tuple[Optional[object], Optional[str], bool, str]:
    provider_name = provider_name or get_default_provider_name()
    profile = get_provider_profile(provider_name)
    return build_llm_provider_from_profile(profile)


def build_llm_provider_from_profile(
    profile: Optional[dict],
) -> tuple[Optional[object], Optional[str], bool, str]:
    if not profile:
        return None, None, False, "找不到目前選擇的 Provider。"
    if OpenAI is None:
        return None, profile.get("model") or None, False, "缺少 openai 套件。"

    provider_name = profile.get("name", "").strip()
    base_url = profile.get("base_url", "").strip()
    api_key = profile.get("api_key", "").strip()
    if not api_key and provider_name:
        # 若使用者未重新輸入，沿用 Windows Credential Manager 中已儲存的 API Key。
        api_key = read_api_key_from_credential_manager(provider_name) or ""
    model = profile.get("model", "").strip()

    missing_items = []
    if is_placeholder(base_url):
        missing_items.append("Base URL")
    if not api_key:
        missing_items.append("API Key")
    if not model:
        missing_items.append("Model")

    if missing_items:
        return None, model or None, False, f"缺少設定：{', '.join(missing_items)}"

    try:
        client = OpenAI(api_key=api_key, base_url=base_url)
    except Exception as exc:
        return None, model or None, False, f"建立 OpenAI client 失敗：{exc}"

    return client, model or None, True, "已就緒"


def fetch_remote_models(profile: Optional[dict]) -> tuple[list[str], str]:
    client, _, _, detail = build_llm_provider_from_profile(profile)
    if client is None:
        return [], detail

    try:
        models_page = client.models.list(timeout=15)
        model_ids = sorted(
            {
                getattr(model, "id", "").strip()
                for model in models_page
                if getattr(model, "id", "").strip()
            }
        )
    except Exception as exc:
        status_code = getattr(exc, "status_code", None)
        if status_code == 404:
            return [], "此服務未提供 /models 端點，請手動輸入 Model。"
        return [], f"讀取模型失敗：{exc}"

    if not model_ids:
        return [], "服務有回應，但沒有列出可用模型。"

    return model_ids, f"已載入 {len(model_ids)} 個模型。"


def create_markitdown_engine(
    provider_name: Optional[str] = None,
    use_ocr: bool = True,
) -> tuple[Optional[object], dict]:
    provider_name = provider_name or get_default_provider_name()
    profile = get_provider_profile(provider_name) or {}
    default_model = profile.get("model", "") or "未設定"
    plugin_ready = has_ocr_plugin()

    if MarkItDown is None:
        return None, {
            "provider_name": provider_name,
            "model": default_model,
            "vision_ready": False,
            "plugin_ready": plugin_ready,
            "used_ocr": False,
            "detail": f"MarkItDown 尚未安裝：{MARKITDOWN_IMPORT_ERROR}",
        }

    if not use_ocr:
        try:
            engine = MarkItDown(enable_plugins=False)
        except TypeError:
            engine = MarkItDown()
        return engine, {
            "provider_name": provider_name,
            "model": default_model,
            "vision_ready": False,
            "plugin_ready": plugin_ready,
            "used_ocr": False,
            "detail": "快速模式：只做文字與結構轉換，不啟用 OCR / Vision。",
        }

    client, llm_model, vision_ready, detail = build_llm_provider(provider_name)
    if not plugin_ready or client is None or not llm_model:
        try:
            engine = MarkItDown(enable_plugins=False)
        except TypeError:
            engine = MarkItDown()
        fallback_reason = "未設定完整 API，OCR / Vision 已停用，會退回快速模式。"
        if detail:
            fallback_reason = f"{detail}；OCR / Vision 已停用，會退回快速模式。"
        return engine, {
            "provider_name": provider_name,
            "model": default_model,
            "vision_ready": False,
            "plugin_ready": plugin_ready,
            "used_ocr": False,
            "detail": fallback_reason,
        }

    try:
        engine = MarkItDown(enable_plugins=True, llm_client=client, llm_model=llm_model)
    except TypeError:
        try:
            engine = MarkItDown(enable_plugins=False)
        except TypeError:
            engine = MarkItDown()
        return engine, {
            "provider_name": provider_name,
            "model": llm_model,
            "vision_ready": False,
            "plugin_ready": plugin_ready,
            "used_ocr": False,
            "detail": "目前安裝的 MarkItDown 版本不支援 llm_client / llm_model，已退回快速模式。",
        }

    return engine, {
        "provider_name": provider_name,
        "model": llm_model,
        "vision_ready": vision_ready,
        "plugin_ready": plugin_ready,
        "used_ocr": True,
        "detail": detail,
    }


def convert_with_markitdown_file(
    file_path: str,
    provider_name: str,
    use_ocr: bool,
) -> tuple[str, str, dict]:
    markitdown_engine, provider_meta = create_markitdown_engine(provider_name, use_ocr=use_ocr)
    if markitdown_engine is None:
        raise RuntimeError(provider_meta.get("detail", "MarkItDown 初始化失敗。"))

    result = markitdown_engine.convert(file_path)
    markdown_content = getattr(result, "text_content", None)
    if markdown_content is None:
        raise RuntimeError("MarkItDown 沒有回傳 text_content，請確認版本與外掛是否正常。")

    output_dir = os.path.abspath(os.path.dirname(file_path))
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.md")

    with open(output_path, "w", encoding="utf-8") as markdown_file:
        markdown_file.write(markdown_content)

    return output_path, output_dir, provider_meta


class ConversionWorker(QThread):
    finished_job = pyqtSignal(dict)

    def __init__(self, job: dict, parent=None):
        super().__init__(parent)
        self.job = job

    def run(self):
        file_path = self.job["file_path"]
        engine = resolve_engine(file_path)
        file_name = os.path.basename(file_path)

        try:
            if engine == "docx" and self.job["docx_mode"] == "mammoth":
                engine_label = "Mammoth"
                converter = DocxToMarkdownConverter()
                output_path, output_dir = converter.convert(file_path)
                provider_meta = {}
            else:
                preferred_ocr = True if engine == "docx" else self.job["ocr_mode"] == "ocr"
                output_path, output_dir, provider_meta = convert_with_markitdown_file(
                    file_path=file_path,
                    provider_name=self.job["provider_name"],
                    use_ocr=preferred_ocr,
                )
                engine_label = "MarkItDown OCR" if provider_meta.get("used_ocr") else "MarkItDown 快速模式"

            self.finished_job.emit(
                {
                    "ok": True,
                    "file_name": file_name,
                    "engine_label": engine_label,
                    "output_path": output_path,
                    "output_dir": output_dir,
                    "detail": provider_meta.get("detail", ""),
                }
            )
        except Exception as exc:
            if engine == "docx" and self.job["docx_mode"] == "mammoth":
                engine_label = "Mammoth"
            else:
                engine_label = "MarkItDown OCR" if self.job["ocr_mode"] == "ocr" else "MarkItDown 快速模式"

            self.finished_job.emit(
                {
                    "ok": False,
                    "file_name": file_name,
                    "engine_label": engine_label,
                    "error": str(exc),
                }
            )


class ProviderEditor(QFrame):
    def __init__(self, profile: Optional[dict] = None, parent=None):
        super().__init__(parent)
        self.profile = profile or {}
        self.remove_requested = None
        self.setObjectName("ProviderCard")
        self.init_ui()

    def init_ui(self):
        self.setStyleSheet(
            """
            QFrame#ProviderCard {
                background: #1c2027;
                border: 1px solid #313845;
                border-radius: 14px;
            }
            QLabel {
                color: #e8eaed;
            }
            QLineEdit, QComboBox {
                background: #11141a;
                color: #ffffff;
                border: 1px solid #465066;
                border-radius: 8px;
                padding: 8px 10px;
                selection-background-color: #7c4dff;
            }
            QLineEdit:focus, QComboBox:focus {
                border: 1px solid #8ab4f8;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 28px;
                border-left: 1px solid #465066;
                background: #18202d;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }
            QComboBox::down-arrow {
                width: 0px;
                height: 0px;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 7px solid #d7def2;
            }
            QPushButton {
                background: #2d6cdf;
                color: white;
                border: 0;
                border-radius: 8px;
                padding: 8px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #4582f0;
            }
            QPushButton#DangerButton {
                background: #b3261e;
            }
            QPushButton#DangerButton:hover {
                background: #d93025;
            }
            """
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setSpacing(8)

        self.title_label = QLabel("Provider")
        self.title_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #f8f9fa;")
        header.addWidget(self.title_label)
        header.addStretch(1)

        self.load_models_button = QPushButton("讀取模型")
        self.load_models_button.clicked.connect(self.load_models)
        header.addWidget(self.load_models_button)

        self.test_button = QPushButton("測試連線")
        self.test_button.clicked.connect(self.test_connection)
        header.addWidget(self.test_button)

        self.remove_button = QPushButton("刪除")
        self.remove_button.setObjectName("DangerButton")
        self.remove_button.clicked.connect(self.request_remove)
        header.addWidget(self.remove_button)

        layout.addLayout(header)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        form.setFormAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(10)

        self.name_input = QLineEdit(self.profile.get("name", ""))
        self.name_input.setPlaceholderText("例如：OpenAI、Anthropic、OpenRouter")
        self.name_input.textChanged.connect(self.on_name_changed)
        form.addRow("Service Provider", self.name_input)

        self.model_input = QComboBox()
        self.model_input.setEditable(True)
        self.model_input.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.model_input.setMaxVisibleItems(18)
        if self.profile.get("model"):
            self.model_input.addItem(self.profile["model"])
            self.model_input.setCurrentText(self.profile["model"])
        self.model_input.lineEdit().setPlaceholderText("例如：gpt-4o-mini")
        form.addRow("Model", self.model_input)

        self.model_hint_label = QLabel()
        self.model_hint_label.setWordWrap(True)
        self.model_hint_label.setStyleSheet("font-size: 11px; color: #aab7d5;")
        form.addRow("", self.model_hint_label)

        self.base_url_input = QLineEdit(self.profile.get("base_url", ""))
        self.base_url_input.setPlaceholderText("https://your-openai-compatible-endpoint/v1")
        form.addRow("Base URL", self.base_url_input)

        self.api_key_input = QLineEdit()
        self.api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("API Key", self.api_key_input)

        layout.addLayout(form)

        self.status_label = QLabel("尚未測試連線。")
        self.status_label.setWordWrap(True)
        self.status_label.setStyleSheet("font-size: 11px; color: #9db2ce;")
        layout.addWidget(self.status_label)

        self.on_name_changed()

    def on_name_changed(self):
        name = self.name_input.text().strip()
        self.title_label.setText(name or "未命名 Provider")
        self.refresh_hints()
        self.refresh_api_key_hint()

    def get_effective_api_key(self) -> str:
        typed_api_key = self.api_key_input.text().strip()
        if typed_api_key:
            return typed_api_key
        provider_name = self.name_input.text().strip()
        if provider_name:
            stored_api_key = read_api_key_from_credential_manager(provider_name) or ""
            if stored_api_key:
                return stored_api_key
        return self.profile.get("api_key", "").strip()

    def refresh_api_key_hint(self):
        if self.get_effective_api_key():
            self.api_key_input.setPlaceholderText("已儲存 API Key；留白可沿用，輸入新值可覆寫")
        else:
            self.api_key_input.setPlaceholderText("尚未儲存 API Key，請輸入後按 Save")

    def refresh_hints(self):
        starter = get_starter_profile_by_name(self.name_input.text().strip())
        if starter:
            self.model_input.lineEdit().setPlaceholderText(starter.get("model", "例如：gpt-4o-mini"))
            self.base_url_input.setPlaceholderText(
                starter.get("base_url", "https://your-openai-compatible-endpoint/v1")
            )
            self.model_hint_label.setText(
                f"建議預設模型：{starter.get('model', '')}。可直接沿用，也可改成你自己的模型 ID。"
            )
        else:
            self.model_input.lineEdit().setPlaceholderText("例如：gpt-4o-mini")
            self.base_url_input.setPlaceholderText("https://your-openai-compatible-endpoint/v1")
            self.model_hint_label.setText(
                "Model 可手動輸入，也可點「讀取模型」從服務端抓取。不是所有服務都支援 /models。"
            )

    def set_removable(self, removable: bool):
        self.remove_button.setVisible(removable)

    def request_remove(self):
        if callable(self.remove_requested):
            self.remove_requested(self)

    def set_status(self, text: str, color: str):
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"font-size: 11px; color: {color};")
        QApplication.processEvents()

    def current_profile(self) -> dict:
        starter = get_starter_profile_by_name(self.name_input.text().strip()) or {}
        return {
            "name": self.name_input.text().strip(),
            "model": self.model_input.currentText().strip() or starter.get("model", ""),
            "base_url": self.base_url_input.text().strip() or starter.get("base_url", ""),
            # 畫面留白時仍沿用已存的憑證，避免誤判缺少 API Key。
            "api_key": self.get_effective_api_key(),
        }

    def apply_model_choices(self, models: list[str]):
        current_value = self.model_input.currentText().strip()
        self.model_input.blockSignals(True)
        self.model_input.clear()
        self.model_input.addItems(models)
        if current_value and current_value in models:
            self.model_input.setCurrentText(current_value)
        elif current_value:
            self.model_input.setEditText(current_value)
        elif models:
            self.model_input.setCurrentText(models[0])
        self.model_input.blockSignals(False)
        if models:
            self.model_input.showPopup()

    def load_models(self):
        profile = self.current_profile()
        self.set_status("讀取模型中...", "#ffb74d")
        models, detail = fetch_remote_models(profile)
        if models:
            self.apply_model_choices(models)
            self.set_status(f"{detail}，請直接從下拉清單選擇。", "#03dac6")
        elif "未提供 /models" in detail:
            self.set_status(detail, "#ffb74d")
        else:
            self.set_status(detail, "#cf6679")

    def test_connection(self):
        profile = self.current_profile()
        self.set_status("測試連線中...", "#ffb74d")
        client, _, _, detail = build_llm_provider_from_profile(profile)
        if client is None:
            self.set_status(detail, "#cf6679")
            return

        models, model_detail = fetch_remote_models(profile)
        if models:
            self.apply_model_choices(models)
            self.set_status("連線成功，已展開模型清單供選擇。", "#03dac6")
        elif "未提供 /models" in model_detail:
            self.set_status("連線成功，但服務未提供 /models；請手動輸入 Model。", "#ffb74d")
        else:
            self.set_status(model_detail or detail, "#cf6679")

    def build_profile(self) -> dict:
        return self.current_profile()


class SettingsDialog(QDialog):
    def __init__(self, settings: dict, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.provider_editors: list[ProviderEditor] = []
        self.saved_settings: Optional[dict] = None
        self.setWindowTitle("設定")
        self.setMinimumSize(760, 640)
        self.init_ui()

    def init_ui(self):
        self.setStyleSheet(
            """
            QDialog {
                background: #121417;
                color: #eef2f6;
            }
            QLabel {
                color: #eef2f6;
            }
            QScrollArea {
                border: 0;
                background: transparent;
            }
            QWidget#ScrollHost {
                background: transparent;
            }
            QPushButton {
                background: #6200ee;
                color: white;
                border: 0;
                border-radius: 10px;
                padding: 10px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #7c4dff;
            }
            QDialogButtonBox QPushButton {
                min-width: 92px;
            }
            """
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(14)

        intro = QLabel(
            f"目前設定檔：{get_settings_display_path()}\n"
            "即使沒有填 API Key，程式仍可使用 DOCX 的 Mammoth 與 MarkItDown 快速模式；只有 OCR / Vision 功能會停用。\n"
            "你可以自由新增多組 OpenAI-compatible Provider；至少需保留一組。"
        )
        intro.setWordWrap(True)
        intro.setStyleSheet("font-size: 13px; color: #d3dbf7;")
        layout.addWidget(intro)

        security_note = QLabel(
            "API Key 會儲存在 Windows Credential Manager，不會明文寫入 settings.json。"
        )
        security_note.setWordWrap(True)
        security_note.setStyleSheet("font-size: 12px; color: #8ad7c7; font-weight: bold;")
        layout.addWidget(security_note)

        env_note = QLabel(
            ".env 只用來提供 Base URL / 範本；若本機尚未儲存 API Key，請在這裡輸入一次。"
        )
        env_note.setWordWrap(True)
        env_note.setStyleSheet("font-size: 12px; color: #aab7d5;")
        layout.addWidget(env_note)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        self.scroll_host = QWidget()
        self.scroll_host.setObjectName("ScrollHost")
        self.providers_layout = QVBoxLayout(self.scroll_host)
        self.providers_layout.setContentsMargins(0, 0, 0, 0)
        self.providers_layout.setSpacing(12)
        self.providers_layout.addStretch(1)

        scroll_area.setWidget(self.scroll_host)
        layout.addWidget(scroll_area, 1)

        action_row = QHBoxLayout()
        action_row.setSpacing(10)

        self.add_provider_button = QPushButton("新增 Provider")
        self.add_provider_button.clicked.connect(self.add_empty_provider)
        action_row.addWidget(self.add_provider_button)
        action_row.addStretch(1)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.accepted.connect(self.on_save)
        self.button_box.rejected.connect(self.reject)
        action_row.addWidget(self.button_box)
        layout.addLayout(action_row)

        providers = self.settings.get("providers", [])
        if not isinstance(providers, list) or not providers:
            providers = build_starter_provider_profiles()

        for profile in providers:
            if isinstance(profile, dict):
                self.add_provider(profile)

        self.refresh_remove_buttons()

    def add_provider(self, profile: Optional[dict] = None):
        editor = ProviderEditor(profile, self)
        editor.remove_requested = self.remove_provider
        self.provider_editors.append(editor)
        self.providers_layout.insertWidget(self.providers_layout.count() - 1, editor)
        self.refresh_remove_buttons()

    def add_empty_provider(self):
        self.add_provider({"name": "", "model": "", "base_url": "", "api_key": ""})

    def remove_provider(self, editor: ProviderEditor):
        if editor not in self.provider_editors:
            return
        self.provider_editors.remove(editor)
        editor.setParent(None)
        editor.deleteLater()
        self.refresh_remove_buttons()

    def refresh_remove_buttons(self):
        removable = len(self.provider_editors) > 1
        for editor in self.provider_editors:
            editor.set_removable(removable)

    def build_settings(self) -> dict:
        providers: list[dict] = []
        seen_names: set[str] = set()

        for editor in self.provider_editors:
            profile = editor.build_profile()
            if not profile["name"]:
                raise ValueError("每一組 Provider 都必須填寫 Service Provider 名稱。")
            if profile["name"] in seen_names:
                raise ValueError(f"Provider 名稱不可重複：{profile['name']}")
            if not profile["model"]:
                raise ValueError(f"Provider「{profile['name']}」尚未填寫 Model。")
            seen_names.add(profile["name"])
            providers.append(profile)

        if not providers:
            raise ValueError("至少要保留一組 Provider。")

        return {
            "providers": providers,
            "ui": dict(self.settings.get("ui", {})),
        }

    def on_save(self):
        try:
            self.saved_settings = self.build_settings()
        except ValueError as exc:
            QMessageBox.warning(self, "設定未完成", str(exc))
            return
        self.accept()


class DragDropArea(QFrame):
    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.default_style = """
            #DragDropArea {
                border: 3px dashed #6200ee;
                border-radius: 22px;
                background: rgba(255, 255, 255, 0.08);
            }
            #DragDropArea:hover {
                background: rgba(255, 255, 255, 0.12);
                border-color: #7c4dff;
            }
        """
        self.active_style = """
            #DragDropArea {
                border: 3px dashed #03dac6;
                border-radius: 22px;
                background: rgba(3, 218, 198, 0.12);
            }
        """

        self.setAcceptDrops(True)
        self.setObjectName("DragDropArea")
        self.setMinimumSize(460, 250)
        self.setStyleSheet(self.default_style)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(12)

        self.label = QLabel("拖曳檔案到這裡開始轉換", self)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("font-size: 20px; color: #ffffff; font-weight: bold;")
        layout.addWidget(self.label)

        self.description = QLabel(
            "支援 .docx / .pdf / .pptx / .xlsx / .csv\n"
            "以及 .txt / .json / .xml / .html / .epub 等格式",
            self,
        )
        self.description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.description.setStyleSheet("font-size: 14px; color: #cfcfcf;")
        layout.addWidget(self.description)

        self.engine_hint = QLabel(
            "DOCX 可選 Mammoth 或 MarkItDown OCR；其他格式走 MarkItDown。",
            self,
        )
        self.engine_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.engine_hint.setStyleSheet("font-size: 13px; color: #9adbd3;")
        layout.addWidget(self.engine_hint)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        supported_files = [
            url.toLocalFile()
            for url in event.mimeData().urls()
            if os.path.exists(url.toLocalFile()) and resolve_engine(url.toLocalFile()) != "unsupported"
        ]

        if supported_files:
            event.acceptProposedAction()
            self.setStyleSheet(self.active_style)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent):
        self.setStyleSheet(self.default_style)
        file_paths = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if os.path.exists(file_path):
                file_paths.append(file_path)
        self.main_window.enqueue_files(file_paths)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.converter = DocxToMarkdownConverter()

        provider_names = get_provider_names()
        if not provider_names:
            USER_SETTINGS["providers"] = build_starter_provider_profiles()
            provider_names = get_provider_names()

        self.active_provider = get_user_ui_value("provider", get_default_provider_name())
        if self.active_provider not in provider_names:
            self.active_provider = provider_names[0]

        self.docx_mode = get_user_ui_value("docx_mode", DEFAULT_DOCX_MODE)
        if self.docx_mode not in {value for _, value in DOCX_MODE_OPTIONS}:
            self.docx_mode = DEFAULT_DOCX_MODE

        self.ocr_mode = get_user_ui_value("ocr_mode", DEFAULT_OCR_MODE)
        if self.ocr_mode not in {value for _, value in OCR_MODE_OPTIONS}:
            self.ocr_mode = DEFAULT_OCR_MODE

        self.markitdown_engine = None
        self.provider_meta = {}
        self.last_output_dir = ""
        self.pending_jobs: list[dict] = []
        self.current_worker: Optional[ConversionWorker] = None
        self.is_processing = False
        self.notice_hide_timer = QTimer(self)
        self.notice_hide_timer.setSingleShot(True)
        self.notice_hide_timer.timeout.connect(self.hide_activity_notice)

        self.init_ui()
        self.rebuild_provider_selector()
        self.refresh_markitdown_engine(show_status=False)

    def init_ui(self):
        self.setWindowTitle(APP_TITLE)
        self.resize(700, 760)
        self.setMinimumSize(640, 700)

        self.setStyleSheet(
            """
            QMainWindow {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                            stop:0 #121212, stop:1 #1c1c1c);
            }
            QLabel {
                color: #e6e6e6;
            }
            QPushButton {
                background-color: #6200ee;
                color: white;
                border-radius: 10px;
                padding: 11px 18px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7c4dff;
            }
            QPushButton:disabled {
                background-color: #333333;
                color: #777777;
            }
            QComboBox {
                background-color: #252525;
                color: #f0f0f0;
                border: 1px solid #444444;
                border-radius: 8px;
                padding: 8px 10px;
                font-size: 13px;
            }
            QComboBox::drop-down {
                border: 0;
                width: 26px;
            }
            """
        )

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(28, 24, 28, 28)
        main_layout.setSpacing(14)

        title = QLabel(APP_TITLE)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #bb86fc;")
        main_layout.addWidget(title)

        subtitle = QLabel("雙引擎 Universal Markdown 轉檔流程")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("font-size: 13px; color: #a5a5a5;")
        main_layout.addWidget(subtitle)

        provider_row = QHBoxLayout()
        provider_row.setSpacing(10)

        provider_title = QLabel("Vision Provider")
        provider_title.setStyleSheet("font-size: 13px; color: #d0d0d0; font-weight: bold;")
        provider_row.addWidget(provider_title)

        self.provider_selector = QComboBox()
        self.provider_selector.currentTextChanged.connect(self.on_provider_changed)
        provider_row.addWidget(self.provider_selector, 1)

        self.settings_btn = QPushButton("設定")
        self.settings_btn.setStyleSheet("padding: 8px 14px; font-size: 13px; border-radius: 8px;")
        self.settings_btn.clicked.connect(self.open_settings_dialog)
        provider_row.addWidget(self.settings_btn)

        main_layout.addLayout(provider_row)

        docx_row = QHBoxLayout()
        docx_row.setSpacing(10)

        docx_title = QLabel("DOCX 模式")
        docx_title.setStyleSheet("font-size: 13px; color: #d0d0d0; font-weight: bold;")
        docx_row.addWidget(docx_title)

        self.docx_mode_selector = QComboBox()
        for label, value in DOCX_MODE_OPTIONS:
            self.docx_mode_selector.addItem(label, value)
        docx_index = self.docx_mode_selector.findData(self.docx_mode)
        self.docx_mode_selector.setCurrentIndex(docx_index if docx_index >= 0 else 0)
        self.docx_mode_selector.currentIndexChanged.connect(self.on_docx_mode_changed)
        docx_row.addWidget(self.docx_mode_selector, 1)

        main_layout.addLayout(docx_row)

        ocr_row = QHBoxLayout()
        ocr_row.setSpacing(10)

        ocr_title = QLabel("其他格式 OCR 模式")
        ocr_title.setStyleSheet("font-size: 13px; color: #d0d0d0; font-weight: bold;")
        ocr_row.addWidget(ocr_title)

        self.ocr_mode_selector = QComboBox()
        for label, value in OCR_MODE_OPTIONS:
            self.ocr_mode_selector.addItem(label, value)
        ocr_index = self.ocr_mode_selector.findData(self.ocr_mode)
        self.ocr_mode_selector.setCurrentIndex(ocr_index if ocr_index >= 0 else 0)
        self.ocr_mode_selector.currentIndexChanged.connect(self.on_ocr_mode_changed)
        ocr_row.addWidget(self.ocr_mode_selector, 1)

        main_layout.addLayout(ocr_row)

        self.drop_area = DragDropArea(self, self)
        main_layout.addWidget(self.drop_area)

        self.provider_panel = QFrame()
        self.provider_panel.setStyleSheet(
            """
            QFrame {
                background: rgba(255, 255, 255, 0.05);
                border: 1px solid rgba(255, 255, 255, 0.08);
                border-radius: 12px;
            }
            """
        )
        provider_panel_layout = QVBoxLayout(self.provider_panel)
        provider_panel_layout.setContentsMargins(14, 12, 14, 12)
        provider_panel_layout.setSpacing(6)

        self.provider_summary_label = QLabel()
        self.provider_summary_label.setWordWrap(True)
        self.provider_summary_label.setStyleSheet("font-size: 12px; color: #8ab4f8; font-weight: bold;")
        provider_panel_layout.addWidget(self.provider_summary_label)

        self.provider_detail_label = QLabel()
        self.provider_detail_label.setWordWrap(True)
        self.provider_detail_label.setStyleSheet("font-size: 11px; color: #d7def2;")
        provider_panel_layout.addWidget(self.provider_detail_label)

        main_layout.addWidget(self.provider_panel)

        self.status_label = QLabel("狀態：等待拖曳檔案")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("font-size: 14px; color: #03dac6;")
        main_layout.addWidget(self.status_label)

        self.activity_panel = QFrame()
        self.activity_panel.setVisible(False)
        self.activity_panel.setStyleSheet(
            """
            QFrame {
                background: rgba(255, 255, 255, 0.06);
                border: 1px solid rgba(255, 255, 255, 0.10);
                border-radius: 12px;
            }
            """
        )
        activity_layout = QVBoxLayout(self.activity_panel)
        activity_layout.setContentsMargins(14, 12, 14, 12)
        activity_layout.setSpacing(6)

        self.activity_label = QLabel("正在轉換中")
        self.activity_label.setStyleSheet("font-size: 13px; color: #ffcf70; font-weight: bold;")
        activity_layout.addWidget(self.activity_label)

        self.activity_detail_label = QLabel("請稍候，程式仍在背景處理，這不是當機。")
        self.activity_detail_label.setWordWrap(True)
        self.activity_detail_label.setStyleSheet("font-size: 11px; color: #d7def2;")
        activity_layout.addWidget(self.activity_detail_label)

        self.activity_progress = QProgressBar()
        self.activity_progress.setRange(0, 0)
        self.activity_progress.setTextVisible(False)
        self.activity_progress.setFixedHeight(10)
        self.activity_progress.setStyleSheet(
            """
            QProgressBar {
                background: rgba(255, 255, 255, 0.08);
                border: 0;
                border-radius: 5px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #ffb74d, stop:1 #03dac6);
                border-radius: 5px;
            }
            """
        )
        activity_layout.addWidget(self.activity_progress)

        main_layout.addWidget(self.activity_panel)

        self.open_folder_btn = QPushButton("開啟輸出資料夾")
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self.open_result_folder)
        main_layout.addWidget(self.open_folder_btn)

        main_layout.addStretch(1)

    def rebuild_provider_selector(self):
        provider_names = get_provider_names()
        if not provider_names:
            USER_SETTINGS["providers"] = build_starter_provider_profiles()
            provider_names = get_provider_names()

        if self.active_provider not in provider_names:
            self.active_provider = provider_names[0]

        self.provider_selector.blockSignals(True)
        self.provider_selector.clear()
        self.provider_selector.addItems(provider_names)
        self.provider_selector.setCurrentText(self.active_provider)
        self.provider_selector.blockSignals(False)

    def set_status(self, text: str, color: str):
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"font-size: 14px; color: {color}; font-weight: bold;")
        QApplication.processEvents()

    def show_activity_notice(self, title: str, detail: str, color: str = "#ffcf70", busy: bool = True):
        self.notice_hide_timer.stop()
        self.activity_label.setText(title)
        self.activity_label.setStyleSheet(f"font-size: 13px; color: {color}; font-weight: bold;")
        self.activity_detail_label.setText(detail)
        self.activity_progress.setVisible(busy)
        self.activity_panel.setVisible(True)
        QApplication.processEvents()

    def hide_activity_notice(self):
        self.activity_panel.setVisible(False)

    def show_completion_notice(self, title: str, detail: str, success: bool = True):
        color = "#03dac6" if success else "#cf6679"
        self.show_activity_notice(title, detail, color=color, busy=False)
        self.notice_hide_timer.start(5000)

    def persist_ui_settings(self):
        USER_SETTINGS.setdefault("ui", {})
        USER_SETTINGS["ui"]["provider"] = self.active_provider
        USER_SETTINGS["ui"]["docx_mode"] = self.docx_mode
        USER_SETTINGS["ui"]["ocr_mode"] = self.ocr_mode
        save_user_settings(USER_SETTINGS)

    def refresh_provider_panel(self):
        profile = get_provider_profile(self.active_provider) or {}
        model_name = self.provider_meta.get("model") or profile.get("model") or "未設定"
        vision_text = "已啟用" if self.provider_meta.get("vision_ready") else "未啟用"
        plugin_text = "已偵測" if self.provider_meta.get("plugin_ready") else "未安裝"
        detail_text = self.provider_meta.get("detail", "")
        docx_mode_text = "Mammoth 保留圖片" if self.docx_mode == "mammoth" else "MarkItDown OCR"
        ocr_mode_text = "快速模式" if self.ocr_mode == "fast" else "OCR 增強模式"

        self.provider_summary_label.setText(
            f"Provider: {self.active_provider} | Model: {model_name} | DOCX: {docx_mode_text} | 其他格式: {ocr_mode_text}"
        )
        self.provider_detail_label.setText(
            f"Vision: {vision_text} | OCR plugin: {plugin_text}\n{detail_text}"
        )

    def refresh_markitdown_engine(self, show_status: bool = True):
        reload_environment_settings()
        self.rebuild_provider_selector()
        self.markitdown_engine, self.provider_meta = create_markitdown_engine(
            self.active_provider,
            use_ocr=self.ocr_mode == "ocr",
        )
        self.refresh_provider_panel()
        if show_status:
            self.set_status(f"已切換 Provider：{self.active_provider}", "#8ab4f8")

    def on_provider_changed(self, provider_name: str):
        if not provider_name:
            return
        self.active_provider = provider_name
        self.persist_ui_settings()
        self.refresh_markitdown_engine(show_status=True)

    def on_docx_mode_changed(self, index: int):
        selected_mode = self.docx_mode_selector.itemData(index)
        self.docx_mode = selected_mode or DEFAULT_DOCX_MODE
        self.persist_ui_settings()
        self.refresh_provider_panel()
        mode_label = "Mammoth 保留圖片" if self.docx_mode == "mammoth" else "MarkItDown OCR"
        self.set_status(f"已切換 DOCX 模式：{mode_label}", "#8ab4f8")

    def on_ocr_mode_changed(self, index: int):
        selected_mode = self.ocr_mode_selector.itemData(index)
        self.ocr_mode = selected_mode or DEFAULT_OCR_MODE
        self.persist_ui_settings()
        self.refresh_markitdown_engine(show_status=False)
        mode_label = "快速模式" if self.ocr_mode == "fast" else "OCR 增強模式"
        self.set_status(f"已切換其他格式 OCR 模式：{mode_label}", "#8ab4f8")

    def open_settings_dialog(self):
        global USER_SETTINGS

        dialog = SettingsDialog(USER_SETTINGS, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        updated_settings = dialog.saved_settings or dialog.build_settings()
        updated_settings.setdefault("ui", {})
        valid_provider_names = get_provider_names(updated_settings)
        if self.active_provider not in valid_provider_names:
            self.active_provider = valid_provider_names[0]

        updated_settings["ui"]["provider"] = self.active_provider
        updated_settings["ui"]["docx_mode"] = self.docx_mode
        updated_settings["ui"]["ocr_mode"] = self.ocr_mode

        try:
            sync_credentials_from_settings(updated_settings, USER_SETTINGS)
            save_user_settings(updated_settings)
        except Exception as exc:
            QMessageBox.critical(self, "設定儲存失敗", str(exc))
            return

        USER_SETTINGS = updated_settings
        self.rebuild_provider_selector()
        self.refresh_markitdown_engine(show_status=False)
        self.set_status("設定已儲存", "#03dac6")

    def enqueue_files(self, file_paths: list[str]):
        supported_paths = [path for path in file_paths if resolve_engine(path) != "unsupported"]
        if not supported_paths:
            self.set_status("找不到支援的檔案格式", "#cf6679")
            return

        for path in supported_paths:
            self.pending_jobs.append(
                {
                    "file_path": path,
                    "provider_name": self.active_provider,
                    "docx_mode": self.docx_mode,
                    "ocr_mode": self.ocr_mode,
                }
            )

        if not self.is_processing:
            self.start_next_job()
        else:
            self.set_status(f"已加入佇列，剩餘 {len(self.pending_jobs)} 個檔案", "#8ab4f8")

    def start_next_job(self):
        if not self.pending_jobs:
            self.is_processing = False
            self.current_worker = None
            self.provider_selector.setEnabled(True)
            self.docx_mode_selector.setEnabled(True)
            self.ocr_mode_selector.setEnabled(True)
            self.settings_btn.setEnabled(True)
            return

        self.is_processing = True
        self.provider_selector.setEnabled(False)
        self.docx_mode_selector.setEnabled(False)
        self.ocr_mode_selector.setEnabled(False)
        self.settings_btn.setEnabled(False)

        job = self.pending_jobs.pop(0)
        file_name = os.path.basename(job["file_path"])
        engine = resolve_engine(job["file_path"])

        if engine == "docx":
            engine_label = "Mammoth" if job["docx_mode"] == "mammoth" else "MarkItDown OCR"
        else:
            engine_label = "MarkItDown OCR" if job["ocr_mode"] == "ocr" else "MarkItDown 快速模式"

        self.open_folder_btn.setVisible(False)
        self.set_status(f"處理中: [{engine_label}] {file_name}...", "#ffb74d")
        if engine_label.endswith("OCR"):
            detail_text = "正在背景轉換並進行 OCR / Vision 解析，時間會比較久，這不是當機。"
        else:
            detail_text = "正在背景轉換檔案，請稍候，完成後會顯示輸出提示。"
        self.show_activity_notice(
            f"正在轉換：{file_name}",
            detail_text,
            color="#ffcf70",
            busy=True,
        )

        self.current_worker = ConversionWorker(job, self)
        self.current_worker.finished_job.connect(self.on_worker_finished)
        self.current_worker.finished.connect(self.current_worker.deleteLater)
        self.current_worker.start()

    def on_worker_finished(self, result: dict):
        if result["ok"]:
            self.last_output_dir = result["output_dir"]
            self.open_folder_btn.setVisible(True)
            detail = result.get("detail", "")
            status_text = f"完成: [{result['engine_label']}] {os.path.basename(result['output_path'])}"
            if detail and "退回快速模式" in detail:
                status_text += "（已退回快速模式）"
            self.set_status(status_text, "#03dac6")
            if self.pending_jobs:
                self.show_completion_notice(
                    "這個檔案已完成，準備處理下一個。",
                    f"已完成 {os.path.basename(result['output_path'])}，佇列中還有 {len(self.pending_jobs)} 個檔案。",
                    success=True,
                )
            else:
                self.show_completion_notice(
                    "轉換完成，可開啟輸出資料夾。",
                    f"已完成 {os.path.basename(result['output_path'])}，可直接點下方按鈕查看輸出結果。",
                    success=True,
                )
        else:
            self.set_status(
                f"失敗: [{result['engine_label']}] {result['error']}",
                "#cf6679",
            )
            if self.pending_jobs:
                self.show_completion_notice(
                    "這個檔案轉換失敗，準備處理下一個。",
                    result["error"],
                    success=False,
                )
            else:
                self.show_completion_notice(
                    "轉換失敗，請查看錯誤訊息。",
                    result["error"],
                    success=False,
                )

        self.start_next_job()

    def open_result_folder(self):
        if self.last_output_dir and os.path.exists(self.last_output_dir):
            if sys.platform == "win32":
                os.startfile(self.last_output_dir)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self.last_output_dir])
            else:
                subprocess.Popen(["xdg-open", self.last_output_dir])


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
