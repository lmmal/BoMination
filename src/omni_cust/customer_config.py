# Customer Configuration for BoMination
# This file can be used to easily configure the application for different deployments

# Default customer to use when none is specified
DEFAULT_CUSTOMER = "generic"

# Customer display names for the UI
CUSTOMER_DISPLAY_NAMES = {
    "farrell": "Farrell",
    "nel": "NEL",
    "primetals": "Primetals",
    "generic": "Generic/Other"
}

# Auto-detection keywords for each customer
# These are used to automatically detect the customer from document content
AUTO_DETECTION_KEYWORDS = {
    "nel": ["NEL HYDROGEN", "PROTON ENERGY", "PROTON P/N"],
    "farrell": ["FARRELL"],
    "primetals": ["PRIMETALS", "PRIMETALS TECHNOLOGIES"],
    # Add more customers here as needed
}

# Customer-specific settings
CUSTOMER_SETTINGS = {
    "farrell": {
        "split_mfg_part": True,
        "header_keywords": ["QTY", "PART", "MFG", "PAF", "DESCRIPTION", "INTERNAL"],
        "reject_keywords": ["PRINTED DRAWING", "REFERENCE ONLY"]
    },
    "nel": {
        "remove_instruction_rows": True,
        "header_keywords": ["ITEM", "QTY", "PART", "DESCRIPTION", "REFERENCE", "MFG"],
        "reject_keywords": ["CUT BACK", "REMOVE", "SHRINK TUBING", "DRAWING NUMBER"]
    },
    "primetals": {
        "split_dual_column": True,
        "header_keywords": ["ITEM", "MFG", "MFGPART", "DESCRIPTION", "QTY"],
        "reject_keywords": ["PRIMETALS TECHNOLOGIES", "CONFIDENTIAL", "PROPRIETARY"]
    },
    "generic": {
        "header_keywords": ["ITEM", "QTY", "PART", "DESCRIPTION", "MANUFACTURER"]
    }
}

# For new company deployments, you can:
# 1. Change DEFAULT_CUSTOMER to the most common customer
# 2. Update CUSTOMER_DISPLAY_NAMES for the UI
# 3. Add new customers to AUTO_DETECTION_KEYWORDS
# 4. Configure CUSTOMER_SETTINGS for each customer
# 5. Add new customer formatters to customer_formatters.py
