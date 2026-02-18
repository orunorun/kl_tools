class KLToolsException(Exception):
    """Base class for all exceptions raised by KL Tools."""
    pass

class PDFProcessingError(KLToolsException):
    """Exception raised for errors in PDF processing."""
    def __init__(self, message):
        super().__init__(message)

class ValidationError(KLToolsException):
    """Exception raised for validation errors."""
    def __init__(self, message):
        super().__init__(message)

class FileOperationError(KLToolsException):
    """Exception raised for file operation errors."""
    def __init__(self, message):
        super().__init__(message)

class ConversionError(KLToolsException):
    """Exception raised for conversion errors."""
    def __init__(self, message):
        super().__init__(message)