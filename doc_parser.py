import json
from typing import Optional, Dict
from io import BytesIO
from bs4 import BeautifulSoup
from .utils import *

class SophonDocParser:
    def __init__(self, config_path: str):
        with open(config_path, 'r') as f:
            config = json.load(f)

        self.vlm_api_url = config['vlm_api_url']
        self.vlm_api_key = config['vlm_api_key']
        self.system_prompts = config['system_prompts']
        self.DEFAULT_PROMPT = config['DEFAULT_PROMPT']
        self.DEFAULT_RECT_PROMPT = config['DEFAULT_RECT_PROMPT']
        self.DEFAULT_ROLE_PROMPT = config['DEFAULT_ROLE_PROMPT']
        self.IMAGE_FACTOR = 8
        self.MIN_PIXELS = 4 * self.IMAGE_FACTOR * self.IMAGE_FACTOR
        self.MAX_PIXELS = 4096 * self.IMAGE_FACTOR * self.IMAGE_FACTOR
        self.MAX_RATIO = 128

    def parser_lv1(input: str, type = "document") -> str:
        """Parse the input base64 string and return markdown in base64."""
        doc_bytes = base64.b64decode(input)

        if type == "document":
            # Check file type based on magic numbers
            if doc_bytes[:4] == b'\xD0\xCF\x11\xE0':  # DOC magic number
                docx_bytes = convert_office_to_xml(doc_bytes, type)
                markdown = convert_docx_to_markdown(docx_bytes)
            
            if doc_bytes[:4] == b'\x50\x4B\x03\x04': # docx magic number
                markdown = convert_docx_to_markdown(doc_bytes)
        
        if type == "spreadsheet":
            
            if doc_bytes[:4] == b'\xD0\xCF\x11\xE0': # xls magic number
                docx_bytes = convert_office_to_xml(doc_bytes, type)
                html = covert_xlsx_to_html(docx_bytes)
                markdown = html
            
            if doc_bytes[:4] == b'\x50\x4B\x03\x04': # xlsx magic number
                html = covert_xlsx_to_html(doc_bytes)
                markdown = html
        
        if type == 'slides':

            if doc_bytes[:4] == b'\xD0\xCF\x11\xE0':
                doc_bytes = convert_office_to_xml(doc_bytes, type)
                markdown = pptx_to_json(doc_bytes)

            if doc_bytes[:4] == b'\x50\x4B\x03\x04':
                markdown = pptx_to_json(doc_bytes)

    def parse_pdf(
            self,
            input: str,
            output_dir: str = './',
            prompt: Optional[Dict] = None,
            api_key: Optional[str] = None,
            base_url: Optional[str] = None,
            model: str = 'gpt-4o',
            verbose: bool = False,
            gpt_worker: int = 1,
    ) -> str:
        doc_bytes = base64.b64decode(input)

        unique_id = str(uuid.uuid4())
        output_dir = f"/tmp/_{unique_id}/"
        temp_pdf_path = f"{output_dir}/temp.pdf"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        with open(temp_pdf_path, 'wb') as f:
            f.write(doc_bytes)

        image_infos = parse_pdf_to_images(temp_pdf_path, output_dir=output_dir)

        result = []
        for index, image_info in enumerate(image_infos, start=1):
            page_image, rect_images = image_info

            messages = [{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "image": page_image,
                    },
                    {"type": "text", "text": self.DEFAULT_ROLE_PROMPT + self.DEFAULT_PROMPT},
                ],
            }]

            if rect_images:
                rect_prompt = self.DEFAULT_RECT_PROMPT + ', '.join(rect_images)
                messages[0]["content"].append({"type": "text", "text": rect_prompt})

            image_inputs, _ = process_vision_info(messages, self.IMAGE_FACTOR, self.MIN_PIXELS, self.MAX_PIXELS, self.MAX_RATIO)
            output_text = callvlm(image_to_base64(image_inputs[0]), self.vlm_api_url, self.vlm_api_key, self.system_prompts)

            response_data = json.loads(output_text)
            content = response_data['choices'][0]['message']['content']

            result.append({
                "page": index,
                "text": content,
                "table": None,
                "figures": [base64.b64encode(rect_image.encode('utf-8')).decode('utf-8') for rect_image in rect_images] if rect_images else None
            })

        json_output = json.dumps(result, ensure_ascii=False)
        context_base64 = base64.b64encode(json_output.encode('utf-8')).decode('utf-8')
        return json.dumps({"context_base64": context_base64}, indent=4, ensure_ascii=False)

    def parser_unified(self, input_base64, type):
        doc_bytes = base64.b64decode(input_base64)

        if type == "document":
            if doc_bytes[:4] == b'\xD0\xCF\x11\xE0':  # DOC magic number
                docx_bytes = convert_office_to_xml(doc_bytes, type)
                document = Document(BytesIO(docx_bytes))
            elif doc_bytes[:4] == b'\x50\x4B\x03\x04':  # DOCX magic number
                document = Document(BytesIO(doc_bytes))
            else:
                raise ValueError("Unsupported document type")

            result = []
            for page_index, page in enumerate(document.paragraphs, start=1):
                result.append({
                    "page": page_index,
                    "text": page.text,
                    "table": None,
                    "figures": None
                })

        elif type == "spreadsheet":
            if doc_bytes[:4] == b'\xD0\xCF\x11\xE0':  # XLS magic number
                docx_bytes = convert_office_to_xml(doc_bytes, type)
                html = covert_xlsx_to_html2(docx_bytes)
            elif doc_bytes[:4] == b'\x50\x4B\x03\x04':  # XLSX magic number
                html = covert_xlsx_to_html2(doc_bytes)
            else:
                raise ValueError("Unsupported spreadsheet type")

            soup = BeautifulSoup(html, 'html.parser')
            tables = soup.find_all('table')
            result = []
            for page_index, table in enumerate(tables, start=1):
                result.append({
                    "page": page_index,
                    "text": None,
                    "table": [str(table)],
                    "figures": None
                })

        elif type == 'slides':
            if doc_bytes[:4] == b'\xD0\xCF\x11\xE0':  # PPT magic number
                doc_bytes = convert_office_to_xml(doc_bytes, type)
                prs = Presentation(BytesIO(doc_bytes))
            elif doc_bytes[:4] == b'\x50\x4B\x03\x04':  # PPTX magic number
                prs = Presentation(BytesIO(doc_bytes))
            else:
                raise ValueError("Unsupported slides type")

            result = []
            for slide_index, slide in enumerate(prs.slides, start=1):
                text = ""
                tables = []
                figures = []

                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            text += ' '.join(run.text for run in paragraph.runs) + "\n"

                    if shape.has_table:
                        table_html = BeautifulSoup("<table></table>", "html.parser")
                        for row in shape.table.rows:
                            tr_tag = table_html.new_tag("tr")
                            for cell in row.cells:
                                td_tag = table_html.new_tag("td")
                                td_tag.string = cell.text
                                tr_tag.append(td_tag)
                            table_html.table.append(tr_tag)
                        tables.append(str(table_html))

                    if shape.shape_type == 13:  # Image type
                        img = shape.image
                        img_bytes = img.blob
                        img_base64 = base64.b64encode(img_bytes).decode('utf-8')
                        figures.append(img_base64)

                result.append({
                    "page": slide_index,
                    "text": text.strip(),
                    "table": tables if tables else None,
                    "figures": figures if figures else None
                })

        else:
            raise ValueError("Unsupported file type")
        
        ctx = json.dumps(result, indent=4, ensure_ascii=False)
        result_ = {
            "context_base64": base64.b64encode(ctx.encode()).decode()
        }
        return json.dumps(result_)