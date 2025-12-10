"""
PDF Operations Service
Handles PDF splitting and merging operations.
"""
import os
from typing import List, Optional, Dict
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import openpyxl


class PDFSplitter:
    """Handles PDF splitting operations."""
    
    def __init__(self, input_pdf_path: str):
        """
        Initialize PDF splitter.
        
        Args:
            input_pdf_path: Path to input PDF file
        """
        self.input_pdf_path = input_pdf_path
        self.reader = PdfReader(input_pdf_path)
        self.total_pages = len(self.reader.pages)
    
    def split_by_count(self, pages_per_split: int, output_dir: str, base_name: str = "split") -> List[str]:
        """
        Split PDF into chunks by page count.
        
        Args:
            pages_per_split: Number of pages per split file
            output_dir: Directory to save split files
            base_name: Base name for output files
            
        Returns:
            List of paths to created split files
        """
        if pages_per_split <= 0:
            raise ValueError("pages_per_split must be greater than 0")
        
        output_files = []
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        current_page = 0
        split_num = 1
        
        while current_page < self.total_pages:
            writer = PdfWriter()
            
            # Add pages to this split
            end_page = min(current_page + pages_per_split, self.total_pages)
            for page_num in range(current_page, end_page):
                writer.add_page(self.reader.pages[page_num])
            
            # Save split file
            output_file = os.path.join(output_dir, f"{base_name}_{split_num}.pdf")
            with open(output_file, 'wb') as f:
                writer.write(f)
            
            output_files.append(output_file)
            print(f"Created split {split_num}: {output_file} (pages {current_page + 1}-{end_page})")
            
            current_page = end_page
            split_num += 1
        
        return output_files
    
    def split_by_names(self, names_file_path: str, pages_per_split: int, output_dir: str) -> List[str]:
        """
        Split PDF and name files according to a name list.
        
        Args:
            names_file_path: Path to Excel or TXT file with names
            pages_per_split: Number of pages per split file
            output_dir: Directory to save split files
            
        Returns:
            List of paths to created split files
        """
        # Read names from file
        names = self._read_names_file(names_file_path)
        
        if not names:
            raise ValueError("Names file is empty or could not be read")
        
        # Calculate required splits
        required_splits = (self.total_pages + pages_per_split - 1) // pages_per_split
        
        if len(names) < required_splits:
            print(f"Warning: Only {len(names)} names provided but {required_splits} splits will be created")
            # Pad with numbered names
            for i in range(len(names), required_splits):
                names.append(f"split_{i + 1}")
        
        output_files = []
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        current_page = 0
        split_num = 0
        
        while current_page < self.total_pages and split_num < len(names):
            writer = PdfWriter()
            
            # Add pages to this split
            end_page = min(current_page + pages_per_split, self.total_pages)
            for page_num in range(current_page, end_page):
                writer.add_page(self.reader.pages[page_num])
            
            # Save split file with custom name
            name = self._sanitize_filename(names[split_num])
            output_file = os.path.join(output_dir, f"{name}.pdf")
            
            # Handle duplicate names
            counter = 1
            while os.path.exists(output_file):
                output_file = os.path.join(output_dir, f"{name}_{counter}.pdf")
                counter += 1
            
            with open(output_file, 'wb') as f:
                writer.write(f)
            
            output_files.append(output_file)
            print(f"Created split '{name}': {output_file} (pages {current_page + 1}-{end_page})")
            
            current_page = end_page
            split_num += 1
        
        return output_files
    
    def _read_names_file(self, file_path: str) -> List[str]:
        """
        Read names from Excel or TXT file.
        
        Args:
            file_path: Path to names file
            
        Returns:
            List of names
        """
        names = []
        ext = Path(file_path).suffix.lower()
        
        if ext in ['.xlsx', '.xls']:
            # Read from Excel (first column)
            try:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                
                for row in ws.iter_rows(min_row=1, values_only=True):
                    if row[0]:  # First column
                        name = str(row[0]).strip()
                        if name:
                            names.append(name)
                
                wb.close()
            except Exception as e:
                print(f"Error reading Excel file: {e}")
                raise
        
        elif ext == '.txt':
            # Read from text file (one name per line)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        name = line.strip()
                        if name:
                            names.append(name)
            except Exception as e:
                print(f"Error reading text file: {e}")
                raise
        
        else:
            raise ValueError(f"Unsupported names file format: {ext}")
        
        return names
    
    def _sanitize_filename(self, name: str) -> str:
        """
        Sanitize filename by removing invalid characters.
        
        Args:
            name: Original filename
            
        Returns:
            Sanitized filename
        """
        # Remove invalid characters
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            name = name.replace(char, '_')
        
        # Limit length
        if len(name) > 100:
            name = name[:100]
        
        return name.strip()


class PDFMerger:
    """Handles PDF merging operations."""
    
    def merge_paired(self, file1_path: str, file2_path: str, output_path: str) -> str:
        """
        Merge two PDFs by interleaving pages (paired merge).
        Page 1 from file1, page 1 from file2, page 2 from file1, page 2 from file2, etc.
        
        Args:
            file1_path: Path to first PDF
            file2_path: Path to second PDF
            output_path: Path for output merged PDF
            
        Returns:
            Path to created merged file
        """
        reader1 = PdfReader(file1_path)
        reader2 = PdfReader(file2_path)
        
        writer = PdfWriter()
        
        pages1 = len(reader1.pages)
        pages2 = len(reader2.pages)
        max_pages = max(pages1, pages2)
        
        print(f"Merging PDFs: {pages1} pages from file1, {pages2} pages from file2")
        
        for i in range(max_pages):
            # Add page from file1 if available
            if i < pages1:
                writer.add_page(reader1.pages[i])
                print(f"Added page {i + 1} from file1")
            
            # Add page from file2 if available
            if i < pages2:
                writer.add_page(reader2.pages[i])
                print(f"Added page {i + 1} from file2")
        
        # Ensure output directory exists
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        
        # Write merged PDF
        with open(output_path, 'wb') as f:
            writer.write(f)
        
        print(f"Created merged PDF: {output_path} ({len(writer.pages)} total pages)")
        return output_path
    
    def merge_sequential(self, file_paths: List[str], output_path: str) -> str:
        """
        Merge multiple PDFs sequentially (all pages from file1, then all from file2, etc.).
        
        Args:
            file_paths: List of PDF file paths to merge
            output_path: Path for output merged PDF
            
        Returns:
            Path to created merged file
        """
        if not file_paths:
            raise ValueError("No files provided for merging")
        
        merger = PdfMerger()
        
        for file_path in file_paths:
            print(f"Appending: {file_path}")
            merger.append(file_path)
        
        # Ensure output directory exists
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        
        # Write merged PDF
        with open(output_path, 'wb') as f:
            merger.write(f)
        
        merger.close()
        print(f"Created merged PDF: {output_path}")
        return output_path
    
    def merge_directory(self, directory_path: str, output_path: str, file_extension: str = '.pdf') -> str:
        """
        Merge all PDF files in a directory sequentially (alphabetically sorted).
        
        Args:
            directory_path: Path to directory containing PDF files
            output_path: Path for output merged PDF
            file_extension: File extension to filter (.pdf, .docx, etc.)
            
        Returns:
            Path to created merged file
        """
        import glob
        
        # Get all files with the specified extension
        pattern = os.path.join(directory_path, f"*{file_extension}")
        file_paths = sorted(glob.glob(pattern))
        
        if not file_paths:
            raise ValueError(f"No {file_extension} files found in {directory_path}")
        
        print(f"Found {len(file_paths)} files to merge in {directory_path}")
        return self.merge_sequential(file_paths, output_path)
    
    def get_page_count(self, pdf_path: str) -> int:
        """
        Get page count of a PDF file.
        
        Args:
            pdf_path: Path to PDF file
            
        Returns:
            Number of pages
        """
        reader = PdfReader(pdf_path)
        return len(reader.pages)
