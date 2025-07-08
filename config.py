"""
Configuration file for Word Document Format Modifier
"""

# Default settings
DEFAULT_FONT_SIZE = 12
DEFAULT_CHINESE_FONT = "微软雅黑"
DEFAULT_ENGLISH_FONT = "Arial"
DEFAULT_CHUNK_SIZE = 100
DEFAULT_THREAD_COUNT = 4

# Performance settings
LARGE_DOCUMENT_THRESHOLD = 50  # paragraphs
MEMORY_LIMIT_MB = 512
MAX_CONCURRENT_DOCUMENTS = 8

# Backup settings
BACKUP_ENABLED = True
BACKUP_RETENTION_DAYS = 7
BACKUP_DIRECTORY = "backups"

# Logging settings
LOG_LEVEL = "INFO"
LOG_FILE = "logs/document_processor.log"
LOG_ROTATION_SIZE = "10MB"
LOG_ROTATION_COUNT = 5

# Font priority mappings
FONT_PRIORITIES = {
    # Chinese fonts (higher number = higher priority)
    "等线": 10,
    "方正黑体": 9,
    "华文细黑": 9,
    "华文黑体": 8,
    "微软雅黑": 8,
    "宋体": 7,
    "黑体": 7,
    "楷体": 6,
    "仿宋": 6,
    
    # English fonts
    "Arial": 10,
    "Times New Roman": 9,
    "Calibri": 9,
    "Helvetica": 8,
    "Georgia": 8,
    "Verdana": 7,
    "Consolas": 7,
}

# Color presets
COLOR_PRESETS = {
    "primary": (0, 0, 0),      # Black
    "secondary": (128, 128, 128),  # Gray
    "accent": (0, 102, 204),    # Blue
    "warning": (255, 165, 0),   # Orange
    "error": (255, 0, 0),       # Red
    "success": (0, 128, 0),     # Green
}

# File patterns
WORD_FILE_PATTERNS = ["*.docx", "*.doc"]
TEMP_FILE_PATTERNS = ["~$*", "*.tmp", "*.temp"]

# Performance optimization flags
ENABLE_CHUNKED_PROCESSING = True
ENABLE_MULTITHREADING = True
ENABLE_MEMORY_MONITORING = True
ENABLE_PROGRESS_TRACKING = True

# Feature flags
ENABLE_FONT_FALLBACK = True
ENABLE_AUTOMATIC_BACKUP = True
ENABLE_ERROR_RECOVERY = True
ENABLE_BATCH_PROCESSING = True