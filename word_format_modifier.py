#!/usr/bin/env python3
"""
Word Document Format Modifier
A comprehensive tool for modifying Word document formats with enhanced features.

Features:
- Extended Chinese and English font support
- Performance optimizations for large documents
- Batch processing capabilities
- Multi-threading support
- Enhanced error handling
- Clean, minimal UI
- Automatic backup functionality
"""

import os
import sys
import time
import shutil
import logging
import argparse
import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass
from enum import Enum

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_COLOR_INDEX
    from docx.oxml.shared import OxmlElement, qn
except ImportError:
    print("Error: python-docx is not installed. Please install it with: pip install python-docx")
    sys.exit(1)

try:
    import psutil
    from tqdm import tqdm
    from colorama import init, Fore, Style
    init()  # Initialize colorama for cross-platform colored output
except ImportError:
    print("Error: Required packages are not installed. Please install with: pip install -r requirements.txt")
    sys.exit(1)


class FontCategory(Enum):
    """Font categories for better organization"""
    CHINESE = "chinese"
    ENGLISH = "english"
    MONOSPACE = "monospace"
    SERIF = "serif"
    SANS_SERIF = "sans_serif"


@dataclass
class FontInfo:
    """Font information structure"""
    name: str
    category: FontCategory
    aliases: List[str]
    priority: int = 0  # Higher priority fonts are preferred


class FontManager:
    """Enhanced font management with comprehensive Chinese font support"""
    
    def __init__(self):
        self.fonts = self._initialize_fonts()
        self.font_aliases = self._build_font_aliases()
        
    def _initialize_fonts(self) -> Dict[str, FontInfo]:
        """Initialize comprehensive font database"""
        fonts = {}
        
        # Chinese fonts with extended support
        chinese_fonts = [
            FontInfo("等线", FontCategory.CHINESE, ["DengXian", "Dengxian", "等线 Light"], 10),
            FontInfo("方正黑体", FontCategory.CHINESE, ["FZHei", "FangZheng", "方正黑体 简体"], 9),
            FontInfo("华文细黑", FontCategory.CHINESE, ["STXihei", "华文细黑", "STHeiti Light"], 9),
            FontInfo("华文黑体", FontCategory.CHINESE, ["STHeiti", "华文黑体", "STHeiti Medium"], 8),
            FontInfo("微软雅黑", FontCategory.CHINESE, ["Microsoft YaHei", "微软雅黑", "Microsoft YaHei UI"], 8),
            FontInfo("宋体", FontCategory.CHINESE, ["SimSun", "NSimSun", "宋体"], 7),
            FontInfo("黑体", FontCategory.CHINESE, ["SimHei", "黑体"], 7),
            FontInfo("楷体", FontCategory.CHINESE, ["KaiTi", "楷体", "KaiTi_GB2312"], 6),
            FontInfo("仿宋", FontCategory.CHINESE, ["FangSong", "FangSong_GB2312", "仿宋"], 6),
            FontInfo("华文宋体", FontCategory.CHINESE, ["STSong", "华文宋体"], 5),
            FontInfo("华文仿宋", FontCategory.CHINESE, ["STFangsong", "华文仿宋"], 5),
            FontInfo("华文楷体", FontCategory.CHINESE, ["STKaiti", "华文楷体"], 5),
            FontInfo("苹方", FontCategory.CHINESE, ["PingFang", "PingFang SC", "苹方"], 4),
            FontInfo("思源黑体", FontCategory.CHINESE, ["Source Han Sans", "思源黑体", "Noto Sans CJK"], 4),
            FontInfo("思源宋体", FontCategory.CHINESE, ["Source Han Serif", "思源宋体", "Noto Serif CJK"], 4),
        ]
        
        # English fonts with variants
        english_fonts = [
            FontInfo("Arial", FontCategory.ENGLISH, ["Arial", "Arial Unicode MS"], 10),
            FontInfo("Times New Roman", FontCategory.ENGLISH, ["Times New Roman", "Times"], 9),
            FontInfo("Calibri", FontCategory.ENGLISH, ["Calibri", "Calibri Light"], 9),
            FontInfo("Helvetica", FontCategory.ENGLISH, ["Helvetica", "Helvetica Neue"], 8),
            FontInfo("Georgia", FontCategory.ENGLISH, ["Georgia", "Georgia Pro"], 8),
            FontInfo("Verdana", FontCategory.ENGLISH, ["Verdana", "Verdana Pro"], 7),
            FontInfo("Tahoma", FontCategory.ENGLISH, ["Tahoma", "Tahoma Bold"], 7),
            FontInfo("Trebuchet MS", FontCategory.ENGLISH, ["Trebuchet MS"], 6),
            FontInfo("Comic Sans MS", FontCategory.ENGLISH, ["Comic Sans MS"], 5),
            FontInfo("Impact", FontCategory.ENGLISH, ["Impact"], 5),
            FontInfo("Courier New", FontCategory.ENGLISH, ["Courier New", "Courier"], 6),
            FontInfo("Consolas", FontCategory.ENGLISH, ["Consolas"], 7),
            FontInfo("Monaco", FontCategory.ENGLISH, ["Monaco"], 6),
        ]
        
        # Build font dictionary
        for font in chinese_fonts + english_fonts:
            fonts[font.name] = font
            
        return fonts
    
    def _build_font_aliases(self) -> Dict[str, str]:
        """Build font alias mapping for improved recognition"""
        aliases = {}
        for font_name, font_info in self.fonts.items():
            # Map each alias to the canonical font name
            for alias in font_info.aliases:
                aliases[alias.lower()] = font_name
            # Also map the canonical name to itself
            aliases[font_name.lower()] = font_name
        return aliases
    
    def normalize_font_name(self, font_name: str) -> Optional[str]:
        """Normalize font name using alias mapping"""
        if not font_name:
            return None
        
        # Direct lookup
        normalized = self.font_aliases.get(font_name.lower())
        if normalized:
            return normalized
        
        # Fuzzy matching for partial matches
        font_lower = font_name.lower()
        for alias, canonical in self.font_aliases.items():
            if font_lower in alias or alias in font_lower:
                return canonical
        
        return font_name  # Return original if no match found
    
    def get_font_suggestions(self, category: FontCategory = None) -> List[str]:
        """Get font suggestions by category"""
        if category is None:
            return list(self.fonts.keys())
        
        return [name for name, info in self.fonts.items() 
                if info.category == category]
    
    def get_best_font_match(self, requested_font: str, category: FontCategory = None) -> str:
        """Get the best font match with fallback"""
        # First try exact match
        normalized = self.normalize_font_name(requested_font)
        if normalized and normalized in self.fonts:
            return normalized
        
        # Fallback to category-specific default
        if category:
            category_fonts = self.get_font_suggestions(category)
            if category_fonts:
                # Return highest priority font in category
                return max(category_fonts, 
                          key=lambda x: self.fonts[x].priority)
        
        # Final fallback
        return "Arial" if category == FontCategory.ENGLISH else "微软雅黑"


class ProgressTracker:
    """Clean progress tracking without excessive emojis"""
    
    def __init__(self, total: int, description: str = "Processing"):
        self.total = total
        self.description = description
        self.progress_bar = tqdm(total=total, desc=description, 
                               bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]')
        self.start_time = time.time()
    
    def update(self, amount: int = 1):
        """Update progress"""
        self.progress_bar.update(amount)
    
    def set_description(self, description: str):
        """Update description"""
        self.progress_bar.set_description(description)
    
    def close(self):
        """Close progress bar"""
        self.progress_bar.close()
        elapsed = time.time() - self.start_time
        print(f"{Fore.GREEN}✓ Completed in {elapsed:.2f}s{Style.RESET_ALL}")


class DocumentProcessor:
    """Enhanced document processor with performance optimizations"""
    
    def __init__(self, font_manager: FontManager):
        self.font_manager = font_manager
        self.logger = self._setup_logger()
        self.stats = {
            'processed_documents': 0,
            'modified_paragraphs': 0,
            'errors': 0,
            'total_processing_time': 0
        }
    
    def _setup_logger(self) -> logging.Logger:
        """Setup logging system"""
        logger = logging.getLogger('DocumentProcessor')
        logger.setLevel(logging.INFO)
        
        # Create file handler
        os.makedirs('logs', exist_ok=True)
        file_handler = logging.FileHandler('logs/document_processor.log')
        file_handler.setLevel(logging.INFO)
        
        # Create console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.WARNING)
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger
    
    def create_backup(self, file_path: Path) -> Path:
        """Create backup of document"""
        backup_dir = file_path.parent / 'backups'
        backup_dir.mkdir(exist_ok=True)
        
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
        backup_path = backup_dir / backup_name
        
        shutil.copy2(file_path, backup_path)
        self.logger.info(f"Created backup: {backup_path}")
        
        return backup_path
    
    def is_file_locked(self, file_path: Path) -> bool:
        """Check if file is locked by another process"""
        try:
            # Try to open file in exclusive mode
            with open(file_path, 'r+b') as f:
                pass
            return False
        except (IOError, OSError):
            return True
    
    def process_document_chunk(self, document: Document, start_idx: int, end_idx: int, 
                             modifications: Dict[str, Any]) -> int:
        """Process a chunk of document paragraphs"""
        modified_count = 0
        
        paragraphs = document.paragraphs[start_idx:end_idx]
        
        for paragraph in paragraphs:
            if self.modify_paragraph(paragraph, modifications):
                modified_count += 1
        
        return modified_count
    
    def modify_paragraph(self, paragraph, modifications: Dict[str, Any]) -> bool:
        """Modify paragraph with enhanced formatting options"""
        modified = False
        
        for run in paragraph.runs:
            # Font modifications
            if 'font_name' in modifications:
                new_font = self.font_manager.get_best_font_match(
                    modifications['font_name']
                )
                if run.font.name != new_font:
                    run.font.name = new_font
                    modified = True
            
            # Font size
            if 'font_size' in modifications:
                target_size = Pt(modifications['font_size'])
                if run.font.size != target_size:
                    run.font.size = target_size
                    modified = True
            
            # Bold
            if 'bold' in modifications:
                if run.font.bold != modifications['bold']:
                    run.font.bold = modifications['bold']
                    modified = True
            
            # Italic
            if 'italic' in modifications:
                if run.font.italic != modifications['italic']:
                    run.font.italic = modifications['italic']
                    modified = True
            
            # Underline
            if 'underline' in modifications:
                if run.font.underline != modifications['underline']:
                    run.font.underline = modifications['underline']
                    modified = True
            
            # Color
            if 'color' in modifications:
                color_value = modifications['color']
                if isinstance(color_value, tuple) and len(color_value) == 3:
                    new_color = RGBColor(*color_value)
                    if run.font.color.rgb != new_color:
                        run.font.color.rgb = new_color
                        modified = True
        
        return modified
    
    def process_document(self, file_path: Path, modifications: Dict[str, Any], 
                        chunk_size: int = 100) -> Dict[str, Any]:
        """Process single document with chunked processing for large files"""
        start_time = time.time()
        result = {
            'file_path': str(file_path),
            'success': False,
            'modified_paragraphs': 0,
            'total_paragraphs': 0,
            'processing_time': 0,
            'error': None
        }
        
        try:
            # Check if file is locked
            if self.is_file_locked(file_path):
                raise Exception(f"File is locked by another process: {file_path}")
            
            # Create backup
            backup_path = self.create_backup(file_path)
            
            # Open document
            document = Document(file_path)
            result['total_paragraphs'] = len(document.paragraphs)
            
            # Process in chunks for large documents
            modified_count = 0
            total_paragraphs = len(document.paragraphs)
            
            # Create progress tracker for large documents
            if total_paragraphs > 50:
                progress = ProgressTracker(total_paragraphs, f"Processing {file_path.name}")
            else:
                progress = None
            
            try:
                for i in range(0, total_paragraphs, chunk_size):
                    end_idx = min(i + chunk_size, total_paragraphs)
                    chunk_modified = self.process_document_chunk(
                        document, i, end_idx, modifications
                    )
                    modified_count += chunk_modified
                    
                    if progress:
                        progress.update(end_idx - i)
                
                # Save document
                document.save(file_path)
                
                result['success'] = True
                result['modified_paragraphs'] = modified_count
                
                self.logger.info(f"Successfully processed {file_path}: {modified_count} paragraphs modified")
                
            finally:
                if progress:
                    progress.close()
        
        except Exception as e:
            result['error'] = str(e)
            self.logger.error(f"Error processing {file_path}: {e}")
            self.stats['errors'] += 1
        
        finally:
            result['processing_time'] = time.time() - start_time
            self.stats['total_processing_time'] += result['processing_time']
        
        return result
    
    def process_documents_batch(self, file_paths: List[Path], modifications: Dict[str, Any], 
                              max_workers: int = 4) -> List[Dict[str, Any]]:
        """Process multiple documents with multi-threading support"""
        results = []
        
        print(f"{Fore.BLUE}Processing {len(file_paths)} documents...{Style.RESET_ALL}")
        
        # Use ThreadPoolExecutor for concurrent processing
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_path = {
                executor.submit(self.process_document, path, modifications): path
                for path in file_paths
            }
            
            # Process results as they complete
            for future in as_completed(future_to_path):
                path = future_to_path[future]
                try:
                    result = future.result()
                    results.append(result)
                    
                    # Print result
                    if result['success']:
                        print(f"{Fore.GREEN}✓{Style.RESET_ALL} {path.name}: "
                              f"{result['modified_paragraphs']}/{result['total_paragraphs']} modified")
                    else:
                        print(f"{Fore.RED}✗{Style.RESET_ALL} {path.name}: {result['error']}")
                        
                except Exception as e:
                    print(f"{Fore.RED}✗{Style.RESET_ALL} {path.name}: {e}")
                    results.append({
                        'file_path': str(path),
                        'success': False,
                        'error': str(e)
                    })
        
        self.stats['processed_documents'] += len(file_paths)
        return results
    
    def get_stats(self) -> Dict[str, Any]:
        """Get processing statistics"""
        return self.stats.copy()


def parse_color(color_str: str) -> Optional[Tuple[int, int, int]]:
    """Parse color string to RGB tuple"""
    if not color_str:
        return None
    
    color_str = color_str.strip().lower()
    
    # Common color names
    color_map = {
        'red': (255, 0, 0),
        'green': (0, 255, 0),
        'blue': (0, 0, 255),
        'black': (0, 0, 0),
        'white': (255, 255, 255),
        'yellow': (255, 255, 0),
        'cyan': (0, 255, 255),
        'magenta': (255, 0, 255),
        'orange': (255, 165, 0),
        'purple': (128, 0, 128),
        'brown': (165, 42, 42),
        'gray': (128, 128, 128),
        'grey': (128, 128, 128),
    }
    
    if color_str in color_map:
        return color_map[color_str]
    
    # Try to parse hex color
    if color_str.startswith('#'):
        try:
            color_str = color_str[1:]
            if len(color_str) == 6:
                return tuple(int(color_str[i:i+2], 16) for i in (0, 2, 4))
        except ValueError:
            pass
    
    # Try to parse RGB tuple
    if ',' in color_str:
        try:
            rgb = tuple(int(x.strip()) for x in color_str.split(','))
            if len(rgb) == 3 and all(0 <= x <= 255 for x in rgb):
                return rgb
        except ValueError:
            pass
    
    return None


def find_word_files(directory: Path, recursive: bool = True) -> List[Path]:
    """Find all Word documents in directory"""
    patterns = ['*.docx', '*.doc']
    files = []
    
    for pattern in patterns:
        if recursive:
            files.extend(directory.rglob(pattern))
        else:
            files.extend(directory.glob(pattern))
    
    # Filter out temporary files
    return [f for f in files if not f.name.startswith('~$')]


def print_statistics(stats: Dict[str, Any]):
    """Print processing statistics"""
    print(f"\n{Fore.CYAN}Processing Statistics:{Style.RESET_ALL}")
    print(f"  Documents processed: {stats['processed_documents']}")
    print(f"  Paragraphs modified: {stats['modified_paragraphs']}")
    print(f"  Errors encountered: {stats['errors']}")
    print(f"  Total processing time: {stats['total_processing_time']:.2f}s")
    
    if stats['processed_documents'] > 0:
        avg_time = stats['total_processing_time'] / stats['processed_documents']
        print(f"  Average time per document: {avg_time:.2f}s")


def main():
    """Main function with comprehensive CLI interface"""
    parser = argparse.ArgumentParser(
        description='Enhanced Word Document Format Modifier',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  # Modify single document
  %(prog)s document.docx --font "微软雅黑" --size 12 --bold

  # Batch process directory
  %(prog)s /path/to/docs/ --font "等线" --size 14 --recursive

  # Advanced formatting
  %(prog)s document.docx --font "Arial" --size 11 --color "red" --underline
        '''
    )
    
    # Input arguments
    parser.add_argument('input', nargs='?', help='Input file or directory path')
    parser.add_argument('--recursive', '-r', action='store_true',
                       help='Process directories recursively')
    
    # Font modifications
    parser.add_argument('--font', '-f', help='Font name (supports Chinese fonts)')
    parser.add_argument('--size', '-s', type=int, help='Font size in points')
    parser.add_argument('--bold', '-b', action='store_true', help='Apply bold formatting')
    parser.add_argument('--italic', '-i', action='store_true', help='Apply italic formatting')
    parser.add_argument('--underline', '-u', action='store_true', help='Apply underline formatting')
    parser.add_argument('--color', '-c', help='Font color (name, hex, or RGB)')
    
    # Processing options
    parser.add_argument('--threads', '-t', type=int, default=4,
                       help='Number of threads for batch processing (default: 4)')
    parser.add_argument('--chunk-size', type=int, default=100,
                       help='Chunk size for large document processing (default: 100)')
    
    # Utility options
    parser.add_argument('--list-fonts', action='store_true',
                       help='List available fonts by category')
    parser.add_argument('--no-backup', action='store_true',
                       help='Skip creating backup files')
    
    args = parser.parse_args()
    
    # Initialize components
    font_manager = FontManager()
    processor = DocumentProcessor(font_manager)
    
    # Handle list-fonts option
    if args.list_fonts:
        print(f"{Fore.CYAN}Available Fonts:{Style.RESET_ALL}")
        for category in FontCategory:
            fonts = font_manager.get_font_suggestions(category)
            print(f"\n{category.value.title()} Fonts:")
            for font in fonts:
                print(f"  - {font}")
        return
    
    # Validate input
    if not args.input:
        print(f"{Fore.RED}Error: Input path is required unless using --list-fonts{Style.RESET_ALL}")
        return
        
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"{Fore.RED}Error: Input path does not exist: {input_path}{Style.RESET_ALL}")
        return
    
    # Build modifications dictionary
    modifications = {}
    
    if args.font:
        modifications['font_name'] = args.font
    if args.size:
        modifications['font_size'] = args.size
    if args.bold:
        modifications['bold'] = True
    if args.italic:
        modifications['italic'] = True
    if args.underline:
        modifications['underline'] = True
    if args.color:
        color_rgb = parse_color(args.color)
        if color_rgb:
            modifications['color'] = color_rgb
        else:
            print(f"{Fore.YELLOW}Warning: Invalid color format: {args.color}{Style.RESET_ALL}")
    
    if not modifications:
        print(f"{Fore.YELLOW}Warning: No modifications specified. Use --help for options.{Style.RESET_ALL}")
        return
    
    # Find files to process
    if input_path.is_file():
        if input_path.suffix.lower() in ['.docx', '.doc']:
            files = [input_path]
        else:
            print(f"{Fore.RED}Error: Not a Word document: {input_path}{Style.RESET_ALL}")
            return
    else:
        files = find_word_files(input_path, args.recursive)
        if not files:
            print(f"{Fore.YELLOW}No Word documents found in: {input_path}{Style.RESET_ALL}")
            return
    
    # Show processing summary
    print(f"{Fore.BLUE}Processing Summary:{Style.RESET_ALL}")
    print(f"  Files to process: {len(files)}")
    print(f"  Modifications: {', '.join(modifications.keys())}")
    print(f"  Threads: {args.threads}")
    print(f"  Backup enabled: {not args.no_backup}")
    
    # Process files
    try:
        results = processor.process_documents_batch(
            files, modifications, args.threads
        )
        
        # Print results summary
        successful = sum(1 for r in results if r['success'])
        failed = len(results) - successful
        
        print(f"\n{Fore.GREEN}Processing Complete!{Style.RESET_ALL}")
        print(f"  Successful: {successful}")
        print(f"  Failed: {failed}")
        
        # Print statistics
        print_statistics(processor.get_stats())
        
        # Show failed files
        if failed > 0:
            print(f"\n{Fore.RED}Failed Files:{Style.RESET_ALL}")
            for result in results:
                if not result['success']:
                    print(f"  - {result['file_path']}: {result.get('error', 'Unknown error')}")
    
    except KeyboardInterrupt:
        print(f"\n{Fore.YELLOW}Processing interrupted by user{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}Unexpected error: {e}{Style.RESET_ALL}")


if __name__ == '__main__':
    main()