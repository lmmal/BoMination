from PIL import Image, ImageEnhance

def preprocess_bom_snippet(image_path):
    logging.debug(f"Preprocessing BoM snippet")
    
    # Read the image from the file path using PIL
    pil_image = Image.open(image_path)
    
    # Enhance contrast and brightness for better processing
    enhancer = ImageEnhance.Contrast(pil_image)
    pil_image = enhancer.enhance(2.0)  # Increase contrast
    enhancer = ImageEnhance.Brightness(pil_image)
    pil_image = enhancer.enhance(1.2)  # Slightly increase brightness
    
    # Convert PIL image to OpenCV format
    image = np.array(pil_image)
    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    
    if image is None:
        logging.error(f"Failed to read image from {image_path}")
        return None
    
    # Convert to grayscale
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Adaptive Thresholding for better binarization
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                   cv2.THRESH_BINARY, blockSize=11, C=2)
 
    # Save preprocessed snippet for debugging
    preprocessed_snippet_path = os.path.join(DEBUG_OUTPUT_DIR, f"preprocessed_snippet_{os.path.basename(image_path)}")
    cv2.imwrite(preprocessed_snippet_path, enhanced)
    logging.info(f"Saved preprocessed BoM snippet to {preprocessed_snippet_path}")
    
    return enhanced