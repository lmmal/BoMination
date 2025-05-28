from ultralytics import YOLO
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Paths
yaml_path = r"data.yaml"
weights_path = r"yolov8n.pt"

# Parameters
img_size = 640
batch_size = 16
epochs = 100
project = "BoM Table Recognization"
name = "Table Rec Model"

# Initialize, train model
def train_model():
    # Debug: Verify paths
    if not os.path.exists(yaml_path):
        logging.error(f"YAML file not found at path: {yaml_path}")
        return
    else:
        logging.info(f"YAML file located: {yaml_path}")
    
    if not os.path.exists(weights_path):
        logging.warning(f"Weights file not found locally. YOLO will attempt to download: {weights_path}")
    else:
        logging.info(f"Weights file located: {weights_path}")

    # Initialize model
    try:
        logging.info("Initializing YOLO model...")
        model = YOLO(weights_path)
        logging.info("YOLO model initialized successfully.")
    except Exception as e:
        logging.error(f"Failed to initialize YOLO model: {e}")
        return

    # Start training
    try:
        logging.info("Starting training...")
        model.train(
            data=yaml_path,
            imgsz=img_size,
            batch_size=batch_size,
            epochs=epochs,
            project=project,
            name=name,
        )
        logging.info("Training completed successfully.")
    except Exception as e:
        logging.error(f"Training failed: {e}")
        return

    # Output training logs directory
    training_logs_path = os.path.join("runs", "detect", name)
    if os.path.exists(training_logs_path):
        logging.info(f"Training logs and model outputs are saved at: {training_logs_path}")
    else:
        logging.warning(f"Expected training logs path not found: {training_logs_path}")

if __name__ == "__main__":
    logging.info("Script started.")
    train_model()
    logging.info("Script finished.")
