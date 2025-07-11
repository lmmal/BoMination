# Configuration file for BoM Automation

CONFIG = {
    "paths": {
        "data_dir": "C:/Users/luke.malkasian/Documents/BoM_Automation/Data",
        "models_dir": "C:/Users/luke.malkasian/Documents/BoM_Automation/Models",
        "outputs_dir": "C:/Users/luke.malkasian/Documents/BoM_Automation/Outputs",
        "scripts_dir": "C:/Users/luke.malkasian/Documents/BoM_Automation/Scripts",
        "ocr_weights": "C:/Users/luke.malkasian/Documents/BoM_Automation/Models/openCV/esdr_2x_weights/EDSR_x2.pb",
        "best_price_excel": "C:/Users/luke.malkasian/Documents/BoM_Automation/Data/best_price_test/oemsecrets_test.xlsx"
    },
    "ocr": {
        "tesseract_cmd": "C:/Program Files/Tesseract-OCR/tesseract.exe",
        "psm_mode_first_pass": 7,
        "psm_mode_second_pass": 6,
        "lang": "eng",
        "preserve_spaces": True,
        "recognize_chars": "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-/"
    },
    "image_processing": {
        "dpi": 300,
        "adaptive_thresh_blocksize": 11,
        "adaptive_thresh_C": 2,
        "denoise_strength": 30,
        "dilate_kernel_size": [1, 1],
        "dilate_iterations": 1,
        "erode_kernel_size": [1, 1],
        "erode_iterations": 1
    },
    "cell_detection": {
        "model_weights": "C:/Users/luke.malkasian/Documents/BoM_Automation/Models/cell_detector.pt",
        "conf_threshold": 0.25,
        "iou_threshold": 0.45,
        "min_area": 50
    },
    "export": {
        "excel_output_dir": "C:/Users/luke.malkasian/Documents/BoM_Automation/Outputs/Excel_Files"
    },
    "logging": {
        "level": "INFO",
        "log_file": "C:/Users/luke.malkasian/Documents/BoM_Automation/Outputs/logs/pipeline.log"
    },
    "parallel_processing": {
        "enabled": True,
        "num_workers": 4
    }
}
