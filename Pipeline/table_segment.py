import logging
import os
import cv2
from ultralytics import YOLO
from config import DEBUG_OUTPUT_DIR

class TableSegmenter:
    def __init__(self, model_path):
        self.model = YOLO(model_path)
        # Configure detection thresholds
        self.model.conf = 0.5  # Confidence threshold
        self.model.iou = 0.4  # IoU threshold for NMS

    def process_boxes(self, results, image, cell_bboxes, cell_snippets):
        # Debug: Log the type and content of results
        logging.debug(f"Type of results: {type(results)}")
        logging.debug(f"Content of results: {results}")

        # Check if results is a list and access the first element
        if isinstance(results, list) and len(results) > 0:
            results = results[0]  # Access the first element of the list

        if not hasattr(results, 'boxes'):
            logging.error("Results object has no attribute 'boxes'.")
            return cell_snippets, cell_bboxes

        if not results.boxes:
            logging.warning("No boxes to process.")
            return cell_snippets, cell_bboxes

        for box in results.boxes:
            if box.conf < 0.5:  # Skip low-confidence boxes
                continue

            bbox = box.xyxy.cpu().numpy().astype(int).tolist()  # Convert to list of integers
            if len(bbox) == 1 and isinstance(bbox[0], list):  # Check if bbox is a list within a list
                bbox = bbox[0]  # Strip the outer list
            if len(bbox) == 4:
                x1, y1, x2, y2 = bbox
                cell_bboxes.append(bbox)
                cell_snippet = image[y1:y2, x1:x2]
                cell_snippets.append(cell_snippet)
            else:
                logging.error(f"Invalid bounding box: {bbox}")
        return cell_snippets, cell_bboxes

    def segment_table(self, image_path, bom_bbox):
        logging.debug(f"Segmenting table in image: {image_path}")
        results = self.model(image_path)  # Perform inference on the image
        cell_bboxes = []  # List to store bounding boxes of cells
        cell_snippets = []  # List to store cell snippets
        image = cv2.imread(image_path)  # Read the image using OpenCV

        if image is None:
            logging.error(f"Failed to read image from {image_path}")
            return [], []
        
        cell_snippets, cell_bboxes = self.process_boxes(results, image, cell_bboxes, cell_snippets)  # Process the results
        
        # Save cell snippets to disk
        for i, cell_snippet in enumerate(cell_snippets):
            snippet_path = os.path.join(DEBUG_OUTPUT_DIR, f"cell_snippet_{i}.png")
            if os.path.exists(snippet_path):
                logging.warning(f"Overwriting existing cell snippet: {snippet_path}")
            cv2.imwrite(snippet_path, cell_snippet)
            logging.info(f"Saved cell snippet {i} to {snippet_path}")
        
        return cell_snippets, cell_bboxes