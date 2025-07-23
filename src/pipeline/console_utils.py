"""
Console utility functions for safe printing in Windows executable.
Handles Unicode encoding issues that occur in PyInstaller executables.
"""

import sys
import os

def safe_print(*args, **kwargs):
    """
    Safe print function that handles Unicode encoding issues in Windows executables.
    Falls back to ASCII-safe characters when Unicode fails.
    """
    try:
        # Try normal print first
        print(*args, **kwargs)
    except UnicodeEncodeError:
        # Convert Unicode characters to ASCII equivalents
        safe_args = []
        for arg in args:
            if isinstance(arg, str):
                # Replace common Unicode characters with ASCII equivalents
                safe_arg = (arg
                    .replace('✅', '[OK]')
                    .replace('❌', '[ERROR]')
                    .replace('⚠️', '[WARNING]')
                    .replace('🔧', '[DEBUG]')
                    .replace('📊', '[DATA]')
                    .replace('📝', '[NOTE]')
                    .replace('🎯', '[TARGET]')
                    .replace('🚀', '[START]')
                    .replace('💾', '[SAVE]')
                    .replace('📦', '[PACKAGE]')
                    .replace('🔍', '[SEARCH]')
                    .replace('🎨', '[FORMAT]')
                    .replace('🔗', '[LINK]')
                    .replace('⏳', '[WAIT]')
                    .replace('🔄', '[PROCESS]')
                    .replace('📄', '[FILE]')
                    .replace('📁', '[FOLDER]')
                    .replace('🌐', '[WEB]')
                    .replace('📏', '[SIZE]')
                    .replace('🎉', '[SUCCESS]')
                    .replace('🧹', '[CLEAN]')
                    .replace('📋', '[TABLE]')
                    .replace('💰', '[PRICE]')
                    .replace('🏢', '[COMPANY]')
                )
                safe_args.append(safe_arg)
            else:
                safe_args.append(arg)
        
        try:
            print(*safe_args, **kwargs)
        except UnicodeEncodeError:
            # Last resort: encode to ASCII with error handling
            final_args = []
            for arg in safe_args:
                if isinstance(arg, str):
                    final_args.append(arg.encode('ascii', errors='replace').decode('ascii'))
                else:
                    final_args.append(str(arg))
            print(*final_args, **kwargs)

def is_executable_environment():
    """
    Check if we're running as a PyInstaller executable.
    Returns True if running as .exe, False if running as Python script.
    """
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

def setup_console_encoding():
    """
    Set up console encoding for Windows executable environment.
    """
    if is_executable_environment() and sys.platform.startswith('win'):
        try:
            # Try to set UTF-8 encoding for console
            import codecs
            sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, errors='replace')
            sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, errors='replace')
        except:
            # If that fails, we'll rely on safe_print function
            pass
