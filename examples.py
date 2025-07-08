#!/usr/bin/env python3
"""
Example usage scripts for Word Document Format Modifier
"""

import os
import sys
from pathlib import Path
from docx import Document
from docx.shared import Pt

def create_sample_documents():
    """Create sample Word documents for testing"""
    
    # Create sample documents directory
    sample_dir = Path("sample_documents")
    sample_dir.mkdir(exist_ok=True)
    
    # Sample 1: Chinese content document
    doc1 = Document()
    doc1.add_heading('中文测试文档', 0)
    doc1.add_paragraph('这是一个用于测试中文字体的示例文档。')
    doc1.add_paragraph('本文档包含多个段落，用于测试批量格式化功能。')
    doc1.add_paragraph('支持的中文字体包括：等线、方正黑体、华文细黑、微软雅黑等。')
    doc1.add_paragraph('文档格式化工具可以批量修改字体、字号、粗体、斜体等属性。')
    doc1.save(sample_dir / "chinese_sample.docx")
    
    # Sample 2: English content document
    doc2 = Document()
    doc2.add_heading('English Test Document', 0)
    doc2.add_paragraph('This is a sample document for testing English font formatting.')
    doc2.add_paragraph('The document contains multiple paragraphs to test batch formatting.')
    doc2.add_paragraph('Supported English fonts include: Arial, Times New Roman, Calibri, Helvetica.')
    doc2.add_paragraph('The formatting tool can modify font family, size, bold, italic, and other properties.')
    doc2.save(sample_dir / "english_sample.docx")
    
    # Sample 3: Mixed content document
    doc3 = Document()
    doc3.add_heading('Mixed Content / 混合内容', 0)
    doc3.add_paragraph('This document contains both English and Chinese text.')
    doc3.add_paragraph('这个文档包含英文和中文混合内容。')
    doc3.add_paragraph('Font formatting should work correctly for both languages.')
    doc3.add_paragraph('字体格式化应该对两种语言都有效。')
    doc3.save(sample_dir / "mixed_sample.docx")
    
    # Sample 4: Large document for performance testing
    doc4 = Document()
    doc4.add_heading('Large Document Performance Test', 0)
    for i in range(100):
        doc4.add_paragraph(f'Paragraph {i+1}: This is a test paragraph with some content '
                          f'to simulate a large document. The performance optimization '
                          f'features should handle this efficiently.')
    doc4.save(sample_dir / "large_sample.docx")
    
    print(f"Created sample documents in: {sample_dir}")
    return sample_dir

def run_font_examples():
    """Run example font formatting operations"""
    
    sample_dir = create_sample_documents()
    
    print("\n" + "="*60)
    print("Font Formatting Examples")
    print("="*60)
    
    # Example 1: Chinese font formatting
    print("\n1. Chinese Font Formatting Examples:")
    print("   python word_format_modifier.py sample_documents/chinese_sample.docx --font '等线' --size 14")
    print("   python word_format_modifier.py sample_documents/chinese_sample.docx --font '方正黑体' --size 12 --bold")
    print("   python word_format_modifier.py sample_documents/chinese_sample.docx --font '华文细黑' --size 13")
    
    # Example 2: English font formatting
    print("\n2. English Font Formatting Examples:")
    print("   python word_format_modifier.py sample_documents/english_sample.docx --font 'Arial' --size 12")
    print("   python word_format_modifier.py sample_documents/english_sample.docx --font 'Times New Roman' --size 11 --italic")
    print("   python word_format_modifier.py sample_documents/english_sample.docx --font 'Calibri' --size 11 --bold")
    
    # Example 3: Batch processing
    print("\n3. Batch Processing Examples:")
    print("   python word_format_modifier.py sample_documents/ --font '微软雅黑' --size 12")
    print("   python word_format_modifier.py sample_documents/ --font 'Arial' --size 11 --bold --recursive")
    
    # Example 4: Advanced formatting
    print("\n4. Advanced Formatting Examples:")
    print("   python word_format_modifier.py sample_documents/mixed_sample.docx --font 'Arial' --size 12 --color 'blue'")
    print("   python word_format_modifier.py sample_documents/mixed_sample.docx --font '等线' --size 14 --bold --underline")
    print("   python word_format_modifier.py sample_documents/mixed_sample.docx --font 'Calibri' --size 11 --color '#FF5733'")
    
    # Example 5: Performance optimization
    print("\n5. Performance Optimization Examples:")
    print("   python word_format_modifier.py sample_documents/large_sample.docx --font 'Arial' --size 10 --chunk-size 50")
    print("   python word_format_modifier.py sample_documents/ --font 'Arial' --size 11 --threads 6")
    
    print("\n6. Utility Examples:")
    print("   python word_format_modifier.py --list-fonts")
    print("   python word_format_modifier.py sample_documents/chinese_sample.docx --font 'Arial' --no-backup")

def run_performance_examples():
    """Run performance testing examples"""
    
    print("\n" + "="*60)
    print("Performance Testing Examples")
    print("="*60)
    
    # Create performance test directory
    perf_dir = Path("performance_test")
    perf_dir.mkdir(exist_ok=True)
    
    # Create multiple documents for batch testing
    for i in range(10):
        doc = Document()
        doc.add_heading(f'Performance Test Document {i+1}', 0)
        for j in range(20):
            doc.add_paragraph(f'Document {i+1}, Paragraph {j+1}: This is test content for performance evaluation.')
        doc.save(perf_dir / f"perf_test_{i+1}.docx")
    
    print(f"Created performance test documents in: {perf_dir}")
    print("\nPerformance test commands:")
    print("1. Single-threaded processing:")
    print("   python word_format_modifier.py performance_test/ --font 'Arial' --size 11 --threads 1")
    print("\n2. Multi-threaded processing:")
    print("   python word_format_modifier.py performance_test/ --font 'Arial' --size 11 --threads 4")
    print("\n3. Optimized batch processing:")
    print("   python word_format_modifier.py performance_test/ --font 'Arial' --size 11 --threads 6 --chunk-size 100")

def run_color_examples():
    """Run color formatting examples"""
    
    print("\n" + "="*60)
    print("Color Formatting Examples")
    print("="*60)
    
    # Create color test document
    color_dir = Path("color_test")
    color_dir.mkdir(exist_ok=True)
    
    doc = Document()
    doc.add_heading('Color Formatting Test', 0)
    doc.add_paragraph('This document will be used to test color formatting.')
    doc.add_paragraph('Different colors can be applied to text content.')
    doc.add_paragraph('Colors can be specified as names, hex values, or RGB tuples.')
    doc.save(color_dir / "color_test.docx")
    
    print(f"Created color test document in: {color_dir}")
    print("\nColor formatting commands:")
    print("1. Named colors:")
    print("   python word_format_modifier.py color_test/color_test.docx --color 'red'")
    print("   python word_format_modifier.py color_test/color_test.docx --color 'blue'")
    print("   python word_format_modifier.py color_test/color_test.docx --color 'green'")
    
    print("\n2. Hex colors:")
    print("   python word_format_modifier.py color_test/color_test.docx --color '#FF0000'")
    print("   python word_format_modifier.py color_test/color_test.docx --color '#00FF00'")
    print("   python word_format_modifier.py color_test/color_test.docx --color '#0000FF'")
    
    print("\n3. RGB colors:")
    print("   python word_format_modifier.py color_test/color_test.docx --color '255,0,0'")
    print("   python word_format_modifier.py color_test/color_test.docx --color '0,255,0'")
    print("   python word_format_modifier.py color_test/color_test.docx --color '0,0,255'")

def main():
    """Main function to run all examples"""
    
    print("Word Document Format Modifier - Example Usage")
    print("=" * 60)
    
    # Check if main script exists
    if not Path("word_format_modifier.py").exists():
        print("Error: word_format_modifier.py not found!")
        print("Please ensure you're running this script from the correct directory.")
        sys.exit(1)
    
    # Run examples
    run_font_examples()
    run_performance_examples()
    run_color_examples()
    
    print("\n" + "="*60)
    print("Examples Complete!")
    print("="*60)
    print("\nTo run any of the above commands, copy and paste them into your terminal.")
    print("Make sure you have installed the required dependencies:")
    print("  pip install -r requirements.txt")
    print("\nYou can also run the test suite:")
    print("  python test_word_modifier.py")

if __name__ == "__main__":
    main()