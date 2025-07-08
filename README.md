# Enhanced Word Document Format Modifier

A comprehensive Python tool for batch processing and modifying Word document formats with advanced features including extensive Chinese font support, performance optimizations, and clean user interface.

## Features

### 🎯 Core Functionality
- **Batch Processing**: Process multiple documents simultaneously
- **Multi-threading**: Concurrent processing for improved performance
- **Large Document Support**: Chunked processing for memory efficiency
- **Automatic Backup**: Creates backups before modifications
- **Progress Tracking**: Real-time progress indicators with minimal UI clutter

### 🔤 Enhanced Font Support
- **Extensive Chinese Fonts**: 等线, 方正黑体, 华文细黑, 华文黑体, 微软雅黑, 宋体, 黑体, 楷体, 仿宋, 苹方, 思源黑体, 思源宋体
- **English Font Variants**: Arial, Times New Roman, Calibri, Helvetica, Georgia, Verdana, Consolas
- **Smart Font Matching**: Intelligent font alias recognition and fallback
- **Font Categories**: Organized by Chinese, English, Serif, Sans-serif, Monospace

### 🎨 Advanced Formatting Options
- **Font Properties**: Name, size, bold, italic, underline
- **Color Support**: RGB colors, hex colors, named colors
- **Batch Operations**: Apply consistent formatting across multiple documents
- **Selective Modifications**: Choose specific formatting options to apply

### ⚡ Performance Optimizations
- **Memory Efficient**: Chunked processing for large documents
- **Concurrent Processing**: Multi-threaded batch operations
- **Progress Monitoring**: Real-time progress tracking
- **Resource Management**: Memory usage optimization

### 🛡️ Robust Error Handling
- **File Lock Detection**: Prevents conflicts with open documents
- **Comprehensive Logging**: Detailed error logs and processing statistics
- **Recovery Options**: Automatic backup and restoration capabilities
- **Validation**: Input validation and error reporting

## Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager

### Install Dependencies
```bash
pip install -r requirements.txt
```

### Required Packages
- `python-docx>=0.8.11`: Word document processing
- `colorama>=0.4.6`: Cross-platform colored output
- `tqdm>=4.66.1`: Progress bars
- `psutil>=5.9.0`: System resource monitoring

## Usage

### Basic Usage

#### Single Document Processing
```bash
# Change font to 微软雅黑 with size 14
python word_format_modifier.py document.docx --font "微软雅黑" --size 14

# Apply bold formatting
python word_format_modifier.py document.docx --font "Arial" --size 12 --bold

# Multiple formatting options
python word_format_modifier.py document.docx --font "等线" --size 13 --bold --italic --color "blue"
```

#### Batch Processing
```bash
# Process all Word documents in a directory
python word_format_modifier.py /path/to/documents/ --font "华文细黑" --size 12

# Recursive processing with multiple options
python word_format_modifier.py /path/to/documents/ --font "Calibri" --size 11 --bold --recursive

# Custom thread count for large batches
python word_format_modifier.py /path/to/documents/ --font "Arial" --size 10 --threads 8
```

### Advanced Usage

#### Color Formatting
```bash
# Named colors
python word_format_modifier.py document.docx --color "red"

# Hex colors
python word_format_modifier.py document.docx --color "#FF5733"

# RGB values
python word_format_modifier.py document.docx --color "255,87,51"
```

#### Performance Optimization
```bash
# Custom chunk size for large documents
python word_format_modifier.py large_document.docx --font "Arial" --chunk-size 200

# Disable backup for faster processing
python word_format_modifier.py document.docx --font "Arial" --no-backup
```

#### Font Discovery
```bash
# List all available fonts by category
python word_format_modifier.py --list-fonts
```

### Command Line Options

#### Input Options
- `input`: Input file or directory path
- `--recursive, -r`: Process directories recursively

#### Font Modifications
- `--font, -f`: Font name (supports Chinese fonts)
- `--size, -s`: Font size in points
- `--bold, -b`: Apply bold formatting
- `--italic, -i`: Apply italic formatting
- `--underline, -u`: Apply underline formatting
- `--color, -c`: Font color (name, hex, or RGB)

#### Processing Options
- `--threads, -t`: Number of threads for batch processing (default: 4)
- `--chunk-size`: Chunk size for large document processing (default: 100)

#### Utility Options
- `--list-fonts`: List available fonts by category
- `--no-backup`: Skip creating backup files

## Examples

### Chinese Font Processing
```bash
# Process with popular Chinese fonts
python word_format_modifier.py document.docx --font "等线" --size 14
python word_format_modifier.py document.docx --font "方正黑体" --size 12 --bold
python word_format_modifier.py document.docx --font "华文细黑" --size 11
python word_format_modifier.py document.docx --font "思源黑体" --size 13
```

### Batch Processing Examples
```bash
# Process all documents in current directory
python word_format_modifier.py . --font "微软雅黑" --size 12

# Process with specific formatting
python word_format_modifier.py ./reports/ --font "Arial" --size 11 --bold --recursive

# High-performance batch processing
python word_format_modifier.py ./large_docs/ --font "Calibri" --size 10 --threads 6 --chunk-size 150
```

### Professional Document Formatting
```bash
# Standard business document formatting
python word_format_modifier.py business_docs/ --font "Calibri" --size 11 --recursive

# Academic document formatting
python word_format_modifier.py academic_papers/ --font "Times New Roman" --size 12 --recursive

# Chinese business documents
python word_format_modifier.py chinese_docs/ --font "等线" --size 12 --recursive
```

## Configuration

The tool uses `config.py` for default settings and customization:

```python
# Default settings
DEFAULT_FONT_SIZE = 12
DEFAULT_CHINESE_FONT = "微软雅黑"
DEFAULT_ENGLISH_FONT = "Arial"

# Performance settings
LARGE_DOCUMENT_THRESHOLD = 50
MAX_CONCURRENT_DOCUMENTS = 8

# Backup settings
BACKUP_ENABLED = True
BACKUP_RETENTION_DAYS = 7
```

## Testing

Run the comprehensive test suite:

```bash
python test_word_modifier.py
```

The test suite includes:
- Font management tests
- Document processing tests
- Batch processing tests
- Utility function tests
- Progress tracking tests

## File Structure

```
word-format-modifier/
├── word_format_modifier.py    # Main application
├── config.py                  # Configuration settings
├── test_word_modifier.py      # Test suite
├── requirements.txt           # Dependencies
├── README.md                  # This file
├── logs/                      # Log files (created automatically)
└── backups/                   # Backup files (created automatically)
```

## Supported Fonts

### Chinese Fonts
- 等线 (DengXian) - Modern, clean font
- 方正黑体 (FangZheng Hei) - Bold, professional
- 华文细黑 (STXihei) - Thin, elegant
- 华文黑体 (STHeiti) - Standard black font
- 微软雅黑 (Microsoft YaHei) - Popular system font
- 宋体 (SimSun) - Traditional serif
- 黑体 (SimHei) - Basic black font
- 楷体 (KaiTi) - Handwriting style
- 仿宋 (FangSong) - Traditional style
- 苹方 (PingFang) - Apple system font
- 思源黑体 (Source Han Sans) - Open source
- 思源宋体 (Source Han Serif) - Open source serif

### English Fonts
- Arial - Clean sans-serif
- Times New Roman - Traditional serif
- Calibri - Modern sans-serif
- Helvetica - Classic sans-serif
- Georgia - Web-friendly serif
- Verdana - Screen-optimized
- Consolas - Monospace coding font
- Trebuchet MS - Rounded sans-serif
- Tahoma - Compact sans-serif

## Performance Tips

1. **Large Documents**: Use chunked processing with `--chunk-size` parameter
2. **Batch Processing**: Increase thread count with `--threads` for faster processing
3. **Memory Usage**: Monitor system resources when processing many large files
4. **Backup Management**: Use `--no-backup` for faster processing when backups aren't needed

## Error Handling

The tool includes comprehensive error handling:
- File lock detection prevents conflicts
- Automatic backup creation protects against data loss
- Detailed logging helps troubleshoot issues
- Recovery options available for failed operations

## Logging

All operations are logged to `logs/document_processor.log` with:
- Processing statistics
- Error details
- Performance metrics
- File modification records

## Contributing

Feel free to contribute by:
1. Adding new font support
2. Improving performance optimizations
3. Enhancing error handling
4. Adding new formatting options
5. Improving documentation

## License

This project is open source and available under the MIT License.

## Support

For issues or questions:
1. Check the logs in `logs/document_processor.log`
2. Run the test suite to verify installation
3. Review the examples in this README
4. Check font availability with `--list-fonts`