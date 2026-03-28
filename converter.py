import os
import mammoth
import base64

class DocxToMarkdownConverter:
    def __init__(self):
        self.image_counter = 0
        self.output_dir = ""
        self.resources_dir = ""

    def _image_handler(self, image):
        """
        處理 Word 中的內嵌圖片：
        將圖片直接寫入 Word 檔案所在的資料夾，
        檔名加上 Word 檔名作為前綴以避免衝突。
        """
        self.image_counter += 1
        
        # 決定圖片副檔名
        extension = image.content_type.split("/")[-1]
        if extension == "jpeg":
            extension = "jpg"
        
        # 檔名格式：Word檔名_image_1.png
        image_filename = f"{self.doc_base_name}_image_{self.image_counter}.{extension}"
        image_path = os.path.join(self.output_dir, image_filename)

        # 寫入圖片二進位資料
        with image.open() as image_bytes:
            with open(image_path, "wb") as f:
                f.write(image_bytes.read())

        # 回傳 Markdown 語法需要的路徑 (直接用檔名)
        return {"src": image_filename}

    def convert(self, docx_path):
        """
        執行轉檔主邏輯
        """
        try:
            print(f"[Debug] 開始轉換檔案: {docx_path}")
            self.image_counter = 0
            self.output_dir = os.path.abspath(os.path.dirname(docx_path))
            self.doc_base_name = os.path.splitext(os.path.basename(docx_path))[0]
            
            output_md_path = os.path.join(self.output_dir, f"{self.doc_base_name}.md")
            
            print(f"[Debug] 輸出路徑: {output_md_path}")
            print(f"[Debug] 圖片將儲存於與 Word 同目錄")

            if not os.path.exists(docx_path):
                raise FileNotFoundError(f"找不到檔案: {docx_path}")

            with open(docx_path, "rb") as docx_file:
                # 使用 mammoth 轉換，並掛載自定義的圖片處理器
                result = mammoth.convert_to_markdown(
                    docx_file,
                    convert_image=mammoth.images.inline(self._image_handler)
                )
                
                if result.messages:
                    print(f"[Mammoth Messages] {result.messages}")
                
                markdown_content = result.value
                # 將產出的 Markdown 寫入檔案
                with open(output_md_path, "w", encoding="utf-8") as md_file:
                    md_file.write(markdown_content)
            
            print(f"[Debug] 轉換完成: {output_md_path}")
            return output_md_path, self.output_dir
        except Exception as e:
            print(f"[Error] 轉換過程中發生錯誤: {str(e)}")
            import traceback
            traceback.print_exc()
            raise e
