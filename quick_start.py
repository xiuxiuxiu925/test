#!/usr/bin/env python3
"""
Quick Start Guide for Word Document Format Modifier
"""

import subprocess
import sys
from pathlib import Path

def print_header(title):
    print("\n" + "=" * 60)
    print(f"  {title}")
    print("=" * 60)

def run_command(cmd, description):
    print(f"\n{description}")
    print(f"Command: {cmd}")
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if result.returncode == 0:
        print("✓ Success")
        if result.stdout:
            print(result.stdout)
    else:
        print("✗ Error")
        if result.stderr:
            print(result.stderr)
    return result.returncode == 0

def main():
    print_header("Word Document Format Modifier - Quick Start")
    
    # Check if dependencies are installed
    print("\n1. Checking Dependencies...")
    try:
        import docx
        import tqdm
        import psutil
        import colorama
        print("✓ All dependencies are installed")
    except ImportError as e:
        print(f"✗ Missing dependency: {e}")
        print("Please run: pip install -r requirements.txt")
        return
    
    # Check if main script exists
    if not Path("word_format_modifier.py").exists():
        print("✗ word_format_modifier.py not found")
        print("Please ensure you're running this from the correct directory")
        return
    
    print_header("Quick Commands")
    
    # Basic commands
    print("\n📝 Basic Usage:")
    print("  python word_format_modifier.py document.docx --font 'Arial' --size 12")
    print("  python word_format_modifier.py document.docx --font '微软雅黑' --size 14 --bold")
    
    print("\n📁 Batch Processing:")
    print("  python word_format_modifier.py /path/to/docs/ --font 'Arial' --size 11")
    print("  python word_format_modifier.py /path/to/docs/ --font '等线' --size 12 --recursive")
    
    print("\n🎨 Advanced Formatting:")
    print("  python word_format_modifier.py document.docx --font 'Arial' --color 'red' --bold")
    print("  python word_format_modifier.py document.docx --font 'Arial' --color '#FF5733' --underline")
    
    print("\n⚡ Performance Options:")
    print("  python word_format_modifier.py large_doc.docx --font 'Arial' --chunk-size 200")
    print("  python word_format_modifier.py docs/ --font 'Arial' --threads 8")
    
    print("\n🔧 Utility Commands:")
    print("  python word_format_modifier.py --list-fonts")
    print("  python word_format_modifier.py document.docx --font 'Arial' --no-backup")
    
    print_header("Supported Features")
    
    print("\n🔤 Font Options:")
    print("  Chinese: 等线, 方正黑体, 华文细黑, 华文黑体, 微软雅黑, 宋体, 黑体, 楷体, 仿宋")
    print("  English: Arial, Times New Roman, Calibri, Helvetica, Georgia, Verdana")
    print("  Monospace: Consolas, Courier New, Monaco")
    
    print("\n🎨 Formatting Options:")
    print("  --font FONT_NAME    : Change font family")
    print("  --size SIZE         : Change font size (points)")
    print("  --bold              : Apply bold formatting")
    print("  --italic            : Apply italic formatting")
    print("  --underline         : Apply underline formatting")
    print("  --color COLOR       : Change font color (name/hex/rgb)")
    
    print("\n⚙️  Processing Options:")
    print("  --threads N         : Number of threads (default: 4)")
    print("  --chunk-size N      : Chunk size for large docs (default: 100)")
    print("  --recursive         : Process subdirectories")
    print("  --no-backup         : Skip backup creation")
    
    print_header("Color Examples")
    
    print("\n🎨 Color Formats:")
    print("  Named colors: red, blue, green, black, white, yellow, etc.")
    print("  Hex colors: #FF0000, #00FF00, #0000FF")
    print("  RGB colors: 255,0,0 or '255,0,0'")
    
    print_header("Testing & Validation")
    
    print("\n🧪 Run Tests:")
    print("  python test_word_modifier.py")
    
    print("\n📋 Generate Examples:")
    print("  python examples.py")
    
    print("\n📖 List Available Fonts:")
    print("  python word_format_modifier.py --list-fonts")
    
    print_header("Performance Tips")
    
    print("\n⚡ Performance Optimization:")
    print("  • Use --threads for batch processing")
    print("  • Use --chunk-size for large documents")
    print("  • Use --no-backup for faster processing")
    print("  • Monitor memory usage with large batches")
    
    print("\n💾 Backup & Safety:")
    print("  • Automatic backups in ./backups/ directory")
    print("  • File lock detection prevents conflicts")
    print("  • Detailed logging in ./logs/ directory")
    print("  • Error recovery options available")
    
    print_header("Common Use Cases")
    
    print("\n📄 Document Standardization:")
    print("  python word_format_modifier.py reports/ --font 'Calibri' --size 11 --recursive")
    
    print("\n🌏 Chinese Document Processing:")
    print("  python word_format_modifier.py chinese_docs/ --font '等线' --size 12 --recursive")
    
    print("\n📊 Large Document Processing:")
    print("  python word_format_modifier.py large_report.docx --font 'Arial' --chunk-size 200")
    
    print("\n🔄 Batch Color Formatting:")
    print("  python word_format_modifier.py docs/ --font 'Arial' --color 'blue' --recursive")
    
    print_header("Troubleshooting")
    
    print("\n❓ Common Issues:")
    print("  • File locked: Close document in Word/other applications")
    print("  • Permission denied: Check file/directory permissions")
    print("  • Memory issues: Reduce --threads or --chunk-size")
    print("  • Font not found: Use --list-fonts to see available options")
    
    print("\n📝 Log Files:")
    print("  • Processing logs: ./logs/document_processor.log")
    print("  • Backup location: ./backups/")
    print("  • Error details: Check log files for troubleshooting")
    
    print("\n" + "=" * 60)
    print("  Ready to process your documents!")
    print("=" * 60)

if __name__ == "__main__":
    main()