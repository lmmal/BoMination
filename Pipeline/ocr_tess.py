import logging
import os
import cv2
import pytesseract
import numpy as np
from multiprocessing import Pool
from functools import partial

# Constants (update these as needed)
DEBUG_OUTPUT_DIR = "debug_output"
TESSERACT_CMD = r"C:\Users\luke.malkasian\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
SUPER_RES_MODEL_PATH = r"C:\Users\luke.malkasian\Documents\BoM_Automation\Models\openCV\esdr_2x_weights\EDSR_x2.pb"
PADDING = 10
KERNEL_SIZE = (2, 2)
SHARPEN_KERNEL = [[0, -1, 0], [-1, 5, -1], [0, -1, 0]]

class OCRProcessor:
    def __init__(self):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
        self.sr_model = None
        self.load_super_resolution_model(SUPER_RES_MODEL_PATH)

    def load_super_resolution_model(self, model_path):
        try:
            self.sr_model = cv2.dnn_superres.DnnSuperResImpl_create()
            self.sr_model.readModel(model_path)
            self.sr_model.setModel("edsr", 2)
            logging.info("Loaded 2x EDSR super-resolution model.")
        except Exception as e:
            logging.error(f"Failed to load super-resolution model: {e}")

    def add_padding(self, image, padding=PADDING):
        return cv2.copyMakeBorder(image, padding, padding, padding, padding, cv2.BORDER_CONSTANT, value=[255, 255, 255])

    def super_resolve_image(self, image):
        try:
            if not self.sr_model:
                raise RuntimeError("Super-resolution model not loaded.")
            result = self.sr_model.upsample(image)
            if result.dtype != np.uint8:
                result = result.astype('uint8')
            return result
        except Exception as e:
            logging.warning(f"SR failed: {e} — falling back to resize.")
            return cv2.resize(image, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    def sharpen_image(self, image):
        kernel = np.array(SHARPEN_KERNEL)
        return cv2.filter2D(image, -1, kernel)

    def binarize_image(self, image):
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        return binary

    def apply_morphology(self, image):
        kernel = np.ones(KERNEL_SIZE, np.uint8)
        return cv2.morphologyEx(image, cv2.MORPH_CLOSE, kernel)

    def preprocess_image(self, cell):
        try:
            cell = self.super_resolve_image(cell)
            cell = self.add_padding(cell)
            cell = self.sharpen_image(cell)
            cell = self.binarize_image(cell)
            cell = self.apply_morphology(cell)
            return cell
        except Exception as e:
            logging.error(f"Preprocessing failed: {e}")
            return None

    def tesseract_ocr(self, image, psm=6, lang='eng'):
        config = (
            f'--psm {psm} --oem 3 -l {lang} '
            '-c preserve_interword_spaces=1 '
            '-c tessedit_char_whitelist="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-/"'
        )
        return pytesseract.image_to_string(image, config=config).strip()

    def process_single_cell(self, cell, bbox, preprocessed_dir):
        try:
            preprocessed_cell = self.preprocess_image(cell)
            if preprocessed_cell is None:
                raise ValueError("Preprocessed cell is None")
            path = os.path.join(preprocessed_dir, f"cell_{bbox[1]}.png")
            cv2.imwrite(path, preprocessed_cell)
            text = self.tesseract_ocr(preprocessed_cell, psm=6)
            if not text:
                text = self.tesseract_ocr(preprocessed_cell, psm=7)
            return text, bbox
        except Exception as e:
            logging.error(f"OCR failed for cell {bbox}: {e}")
            return "", bbox

    def perform_ocr(self, cell_snippets, cell_bboxes):
        os.makedirs(DEBUG_OUTPUT_DIR, exist_ok=True)
        pre_dir = os.path.join(DEBUG_OUTPUT_DIR, "preprocessed_cells")
        os.makedirs(pre_dir, exist_ok=True)

        with Pool() as pool:
            results = pool.starmap(
                partial(self.process_single_cell, preprocessed_dir=pre_dir),
                zip(cell_snippets, cell_bboxes)
            )

        data = [text for text, _ in results]
        ocr_mapping = {f"cell_{i}": {"text": text, "bbox": bbox} for i, (text, bbox) in enumerate(results)}

        # Save mapping
        map_path = os.path.join(DEBUG_OUTPUT_DIR, "ocr_mapping.txt")
        try:
            with open(map_path, 'w') as f:
                for key, val in ocr_mapping.items():
                    f.write(f"{key}: {val['text']} (bbox: {val['bbox']})\n")
        except Exception as e:
            logging.error(f"Failed to save OCR mapping: {e}")

        return data