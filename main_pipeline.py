import logging
import os
import cv2
import time

from config import EXCEL_OUTPUT_DIR, DEBUG_OUTPUT_DIR, BOM_REC_MODEL_PATH, TABLE_SEG_MODEL_PATH, OMNI_BASE_COST_SHEET_PATH, OEM_OUTPUT_DIR
from pdf_processor import PDFProcessor
from bom_detector import BoMDetector
from image_preprocessor import ImagePreprocessor
from table_segmenter import TableSegmenter
from ocr_processor import OCRProcessor
from excel_exporter import ExcelExporter
from CostSheet_conv import CostSheetConverter
from oem_best_price import OEMBestPriceAutomator, Info2CostSheet

class BoMPipeline:
    def __init__(self, pdf_path, pages):
        customer_excel_path = os.path.join(EXCEL_OUTPUT_DIR, "customer_BoM.xlsx")
        oem_best_price_path = os.path.join(OEM_OUTPUT_DIR)

        self.pdf_processor = PDFProcessor(pdf_path, pages)
        self.bom_detector = BoMDetector(BOM_REC_MODEL_PATH)
        self.image_preprocessor = ImagePreprocessor()
        self.table_segmenter = TableSegmenter(TABLE_SEG_MODEL_PATH)
        self.ocr_processor = OCRProcessor()
        self.excel_exporter = ExcelExporter(output_dir=EXCEL_OUTPUT_DIR)
        self.cost_sheet_converter = CostSheetConverter(
            customer_excel_path,
            OMNI_BASE_COST_SHEET_PATH,
            os.path.join(EXCEL_OUTPUT_DIR, "converted_cost_sheet.xlsx")
        )
        self.oem_best_price_automator = OEMBestPriceAutomator(oem_best_price_path)
        self.info2costsheet = Info2CostSheet(
            oem_best_price_path,
            os.path.join(EXCEL_OUTPUT_DIR, "final_cost_sheet.xlsx")
        )

    def wait_for_resume_signal(self):
        print("Waiting for user to resume...")
        resume_file = os.path.join(EXCEL_OUTPUT_DIR, "resume.txt")
        while not os.path.exists(resume_file):
            time.sleep(1)
        print("Resume signal received.")
        os.remove(resume_file)

    def run(self):
        images = self.pdf_processor.pdf_to_images()

        for image_path in images:
            logging.info(f"Processing image: {image_path}")

            bom_bbox = self.bom_detector.detect_bom(image_path)
            if not bom_bbox:
                logging.error(f"BoM not found in {image_path}. Saving debug image.")
                continue

            image = cv2.imread(image_path)
            image_height, image_width = image.shape[:2]

            page_num = int(os.path.basename(image_path).split('_')[1].split('.')[0])
            snippet_path = self.pdf_processor.extract_bom_snippet(page_num, bom_bbox, (image_width, image_height))

            preprocessed_snippet = self.image_preprocessor.preprocess_bom_snippet(snippet_path)
            preprocessed_snippet_path = os.path.join(DEBUG_OUTPUT_DIR, f"preprocessed_{os.path.basename(snippet_path)}")
            cv2.imwrite(preprocessed_snippet_path, preprocessed_snippet)

            cell_snippets, cell_bboxes = self.table_segmenter.segment_table(preprocessed_snippet_path, bom_bbox)
            data = self.ocr_processor.perform_ocr(cell_snippets, cell_bboxes)

            mapping_file_path = os.path.join(DEBUG_OUTPUT_DIR, "ocr_mapping.txt")
            excel_path = os.path.join(EXCEL_OUTPUT_DIR, f"BoM_{os.path.basename(image_path).replace('.png', '.xlsx')}")
            self.excel_exporter.export_to_excel(data, excel_path, cell_bboxes, mapping_file_path)
            logging.info(f"Exported BoM to {excel_path}")

            print(f"Please review file: {excel_path} and make edits if necessary.")
            self.wait_for_resume_signal()

            self.cost_sheet_converter.merge_with_omni_format()
            logging.info("Converted customer cost sheet to Omni format")

            self.oem_best_price_automator.automate_bom_tool()
            logging.info("Automated OEM best price retrieval")

            self.info2costsheet.update_omni_cost_sheet()
            logging.info("Updated Omni cost sheet with best prices")
