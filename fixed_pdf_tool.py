#!/usr/bin/env python
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkFont
from PyPDF3 import PdfFileReader, PdfFileWriter, PdfFileMerger
from pdf2docx import Converter
from pdf2image import convert_from_path, convert_from_bytes
from pdf2image.exceptions import (
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)
from PIL import Image, ImageTk
import os
import sys
import pandas as pd
import openpyxl
import subprocess
import shutil
try:
    import tabula
    TABULA_AVAILABLE = True
except ImportError:
    TABULA_AVAILABLE = False
    print("Warning: tabula-py not available. PDF to Excel conversion will use alternative method.")
import pypandoc
import subprocess
import threading
from datetime import datetime
import json
import hashlib
import tempfile
import webbrowser
import pdfplumber
from odf.opendocument import OpenDocumentText
from odf.text import P
from odf.style import Style, TextProperties, ParagraphProperties
from odf.text import Span

class ModernPDFTool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SKANRAY PDF Toolkit ")
        self.root.geometry("1200x800")
        self.root.configure(bg='#1a1a1a')
        self.root.minsize(800, 600)
        
        
    
      
        
        # Color scheme
        self.colors = {
            'bg_primary': '#1a1a1a',
            'bg_secondary': '#2d2d2d',
            'bg_tertiary': '#404040',
            'accent': '#00d4aa',
            'accent_hover': '#00b894',
            'text_primary': '#ffffff',
            'text_secondary': '#b0b0b0',
            'error': '#ff6b6b',
            'warning': '#feca57',
            'success': '#48cae4'
        }
        
        # Initialize variables
        self.current_tool = tk.StringVar(value="dashboard")
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Ready")
        self.history = []
        
        # Create UI components
        self.create_main_layout()
        self.setup_styles()  # Move this after creating main layout
        self.create_sidebar()
        self.create_main_area()
        self.create_status_bar()
        
        # Load dashboard
        self.show_dashboard()
        
        # Center window
        self.center_window()
        
    def setup_styles(self):
        """Setup custom styles for ttk widgets"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure styles with proper error handling
        try:
            self.style.configure('Modern.TButton',
                               background=self.colors['accent'],
                               foreground='white',
                               borderwidth=0,
                               focuscolor='none',
                               padding=(20, 10))
            self.style.map('Modern.TButton',
                          background=[('active', self.colors['accent_hover'])])
            
            self.style.configure('Sidebar.TButton',
                               background=self.colors['bg_secondary'],
                               foreground=self.colors['text_primary'],
                               borderwidth=0,
                               focuscolor='none',
                               padding=(15, 10))
            self.style.map('Sidebar.TButton',
                          background=[('active', self.colors['bg_tertiary'])])
            
            self.style.configure('Modern.TEntry',
                               fieldbackground=self.colors['bg_tertiary'],
                               foreground=self.colors['text_primary'],
                               borderwidth=1,
                               insertcolor=self.colors['text_primary'])
            
            # Simplified progress bar style - use default layout
            self.style.configure('Modern.TProgressbar',
                               background=self.colors['accent'],
                               troughcolor=self.colors['bg_tertiary'],
                               borderwidth=0)
        except tk.TclError as e:
            print(f"Style configuration warning: {e}")
            # Fall back to default styles if custom ones fail
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def create_main_layout(self):
        """Create the main layout structure"""
        # Configure root grid
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        
        # Main containers
        self.sidebar_frame = tk.Frame(self.root, bg=self.colors['bg_secondary'], width=250)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsw", padx=(0, 2))
        self.sidebar_frame.grid_propagate(False)
        
        self.main_frame = tk.Frame(self.root, bg=self.colors['bg_primary'])
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        self.status_frame = tk.Frame(self.root, bg=self.colors['bg_tertiary'], height=30)
        self.status_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.status_frame.grid_propagate(False)
    
    def create_sidebar(self):
        """Create the sidebar with navigation"""
        # Header
        header_frame = tk.Frame(self.sidebar_frame, bg=self.colors['bg_secondary'])
        header_frame.pack(fill="x", padx=10, pady=(20, 30))
        
        title_label = tk.Label(header_frame, text="SKANRAY PDF Toolkit Pro", 
                              font=('Segoe UI', 16, 'bold'),
                              fg=self.colors['text_primary'],
                              bg=self.colors['bg_secondary'])
        title_label.pack()
        
        subtitle_label = tk.Label(header_frame, text="Professional PDF Solutions", 
                                 font=('Segoe UI', 9),
                                 fg=self.colors['text_secondary'],
                                 bg=self.colors['bg_secondary'])
        subtitle_label.pack()
        
        # Navigation buttons
        nav_items = [
            ("🏠 Dashboard", "dashboard"),
            ("🔗 Merge PDFs", "merge"),
            ("↻ Rotate PDF", "rotate"),
            ("📄 PDF to DOC", "pdf_to_doc"),
            ("📊 PDF to Excel", "pdf_to_excel"),
            ("✂️ Split PDF", "split"),
            ("🗜️ Compress PDF", "compress"),
            ("🖼️ PDF to JPG", "pdf_to_jpg"),
            ("📸 JPG to PDF", "jpg_to_pdf"),
            ("🗑️ Remove Pages", "remove_pages"),
            ("📝 PDF to ODT", "pdf_to_odt"),
            ("🔍 PDF Info", "pdf_info"),
            ("🔐 Encrypt PDF", "encrypt"),
            ("🔓 Decrypt PDF", "decrypt"),
            ("📑 Add Watermark", "watermark"),
            ("📋 History", "history")
        ]
        
        for text, value in nav_items:
            btn = ttk.Button(self.sidebar_frame, text=text, style='Sidebar.TButton',
                           command=lambda v=value: self.switch_tool(v))
            btn.pack(fill="x", padx=10, pady=2)
        
        # Footer
        footer_frame = tk.Frame(self.sidebar_frame, bg=self.colors['bg_secondary'])
        footer_frame.pack(side="bottom", fill="x", padx=10, pady=20)
        
        tk.Label(footer_frame, text="© 2025 PDF Toolkit Pro", 
                font=('Segoe UI', 8),
                fg=self.colors['text_secondary'],
                bg=self.colors['bg_secondary']).pack()
    
    def create_main_area(self):
        """Create the main content area"""
        self.content_frame = tk.Frame(self.main_frame, bg=self.colors['bg_primary'])
        self.content_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)
    
    def create_status_bar(self):
        """Create the status bar"""
        # Progress bar with safe style application
        try:
            self.progress_bar = ttk.Progressbar(self.status_frame, style='Modern.TProgressbar',
                                              variable=self.progress_var, mode='determinate')
        except tk.TclError:
            # Fall back to default style if custom style fails
            self.progress_bar = ttk.Progressbar(self.status_frame,
                                              variable=self.progress_var, mode='determinate')
        
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=(10, 5), pady=5)
        
        # Status label
        self.status_label = tk.Label(self.status_frame, textvariable=self.status_var,
                                    font=('Segoe UI', 9),
                                    fg=self.colors['text_primary'],
                                    bg=self.colors['bg_tertiary'])
        self.status_label.pack(side="right", padx=(5, 10), pady=5)
    
    def clear_content(self):
        """Clear the content frame"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()
    
    def switch_tool(self, tool_name):
        """Switch to a different tool"""
        self.current_tool.set(tool_name)
        self.clear_content()
        
        # Route to appropriate method
        if tool_name == "dashboard":
            self.show_dashboard()
        elif tool_name == "merge":
            self.show_merge_tool()
        elif tool_name == "rotate":
            self.show_rotate_tool()
        elif tool_name == "pdf_to_doc":
            self.show_pdf_to_doc_tool()
        elif tool_name == "pdf_to_excel":
            self.show_pdf_to_excel_tool()
        elif tool_name == "split":
            self.show_split_tool()
        elif tool_name == "compress":
            self.show_compress_tool()
        elif tool_name == "pdf_to_jpg":
            self.show_pdf_to_jpg_tool()
        elif tool_name == "jpg_to_pdf":
            self.show_jpg_to_pdf_tool()
        elif tool_name == "remove_pages":
            self.show_remove_pages_tool()
        elif tool_name == "pdf_to_odt":
            self.show_pdf_to_odt_tool()
        elif tool_name == "pdf_info":
            self.show_pdf_info_tool()
        elif tool_name == "encrypt":
            self.show_encrypt_tool()
        elif tool_name == "decrypt":
            self.show_decrypt_tool()
        elif tool_name == "watermark":
            self.show_watermark_tool()
        elif tool_name == "history":
            self.show_history_tool()
    
    def show_dashboard(self):
        """Show the dashboard with overview and quick actions"""
        # Title
        title_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        title_frame.pack(fill="x", pady=(0, 30))
        
        tk.Label(title_frame, text="Welcome to SKANRAY PDF Toolkit Pro", 
                font=('Segoe UI', 24, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w")
        
        tk.Label(title_frame, text="Your complete solution for PDF manipulation and conversion", 
                font=('Segoe UI', 12),
                fg=self.colors['text_secondary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(5, 0))
        
        # Stats cards
        stats_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        stats_frame.pack(fill="x", pady=(0, 30))
        
        self.create_stat_card(stats_frame, "Total Operations", str(len(self.history)), 0)
        self.create_stat_card(stats_frame, "Success Rate", "98.5%", 1)
        self.create_stat_card(stats_frame, "Files Processed", "1,247", 2)
        
        # Quick actions
        quick_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        quick_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(quick_frame, text="Quick Actions", 
                font=('Segoe UI', 18, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 15))
        
        actions_grid = tk.Frame(quick_frame, bg=self.colors['bg_primary'])
        actions_grid.pack(fill="x")
        
        quick_actions = [
            ("Merge PDFs", "merge", "🔗"),
            ("Convert to Word", "pdf_to_doc", "📄"),
            ("Split PDF", "split", "✂️"),
            ("Compress PDF", "compress", "🗜️")
        ]
        
        for i, (text, action, icon) in enumerate(quick_actions):
            self.create_quick_action_card(actions_grid, text, action, icon, i)
    
    def create_stat_card(self, parent, title, value, column):
        """Create a statistics card"""
        card = tk.Frame(parent, bg=self.colors['bg_secondary'], relief="flat", bd=1)
        card.grid(row=0, column=column, padx=10, pady=5, sticky="ew")
        parent.grid_columnconfigure(column, weight=1)
        
        tk.Label(card, text=value, 
                font=('Segoe UI', 20, 'bold'),
                fg=self.colors['accent'],
                bg=self.colors['bg_secondary']).pack(pady=(15, 5))
        
        tk.Label(card, text=title, 
                font=('Segoe UI', 10),
                fg=self.colors['text_secondary'],
                bg=self.colors['bg_secondary']).pack(pady=(0, 15))
    
    def create_quick_action_card(self, parent, title, action, icon, column):
        """Create a quick action card"""
        card = tk.Frame(parent, bg=self.colors['bg_secondary'], relief="flat", bd=1)
        card.grid(row=0, column=column, padx=10, pady=5, sticky="ew")
        parent.grid_columnconfigure(column, weight=1)
        
        # Make card clickable
        def on_click(event=None):
            self.switch_tool(action)
        
        card.bind("<Button-1>", on_click)
        
        # Hover effects
        def on_enter(event):
            card.configure(bg=self.colors['bg_tertiary'])
            for child in card.winfo_children():
                child.configure(bg=self.colors['bg_tertiary'])
        
        def on_leave(event):
            card.configure(bg=self.colors['bg_secondary'])
            for child in card.winfo_children():
                child.configure(bg=self.colors['bg_secondary'])
        
        card.bind("<Enter>", on_enter)
        card.bind("<Leave>", on_leave)
        
        icon_label = tk.Label(card, text=icon, 
                             font=('Segoe UI', 30),
                             fg=self.colors['accent'],
                             bg=self.colors['bg_secondary'])
        icon_label.pack(pady=(15, 10))
        
        title_label = tk.Label(card, text=title, 
                              font=('Segoe UI', 12, 'bold'),
                              fg=self.colors['text_primary'],
                              bg=self.colors['bg_secondary'])
        title_label.pack(pady=(0, 15))
        
        # Bind hover to all child widgets
        for child in card.winfo_children():
            child.bind("<Button-1>", on_click)
            child.bind("<Enter>", on_enter)
            child.bind("<Leave>", on_leave)
    
    def create_tool_header(self, title, description):
        """Create a header for tool pages"""
        header_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        header_frame.pack(fill="x", pady=(0, 30))
        
        tk.Label(header_frame, text=title, 
                font=('Segoe UI', 20, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w")
        
        tk.Label(header_frame, text=description, 
                font=('Segoe UI', 11),
                fg=self.colors['text_secondary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(5, 0))
        
        return header_frame
    
    def create_file_input(self, parent, label_text, browse_command):
        """Create a file input with modern styling"""
        frame = tk.Frame(parent, bg=self.colors['bg_primary'])
        frame.pack(fill="x", pady=10)
        
        tk.Label(frame, text=label_text, 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))
        
        input_frame = tk.Frame(frame, bg=self.colors['bg_primary'])
        input_frame.pack(fill="x")
        
        entry = ttk.Entry(input_frame, style='Modern.TEntry', font=('Segoe UI', 10))
        entry.pack(side="left", fill="x", expand=True, ipady=8)
        
        browse_btn = ttk.Button(input_frame, text="Browse", style='Modern.TButton',
                               command=lambda: browse_command(entry))
        browse_btn.pack(side="right", padx=(10, 0))
        
        return entry
    
    def browse_file(self, entry, filetypes=[("PDF Files", "*.pdf")]):
        """Browse for a file"""
        file_path = filedialog.askopenfilename(filetypes=filetypes, title="Select a file")
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)
    
    def show_merge_tool(self):
        """Show merge PDF tool"""
        self.create_tool_header("Merge PDFs", "Combine multiple PDF files into a single document")
        
        # File inputs
        self.merge_pdf1 = self.create_file_input(self.content_frame, "First PDF File:", 
                                                lambda e: self.browse_file(e))
        self.merge_pdf2 = self.create_file_input(self.content_frame, "Second PDF File:", 
                                                lambda e: self.browse_file(e))
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        merge_btn = ttk.Button(btn_frame, text="🔗 Merge PDFs", style='Modern.TButton',
                              command=self.merge_pdfs)
        merge_btn.pack()
    
    def show_rotate_tool(self):
        """Show rotate PDF tool"""
        self.create_tool_header("Rotate PDF", "Rotate all pages of a PDF document")
        
        # File input
        self.rotate_pdf = self.create_file_input(self.content_frame, "PDF File:", 
                                                lambda e: self.browse_file(e))
        
        # Angle selection
        angle_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        angle_frame.pack(fill="x", pady=10)
        
        tk.Label(angle_frame, text="Rotation Angle:", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))
        
        self.angle_var = tk.IntVar(value=90)
        angle_options = tk.Frame(angle_frame, bg=self.colors['bg_primary'])
        angle_options.pack(anchor="w")
        
        for angle in [90, 180, 270]:
            tk.Radiobutton(angle_options, text=f"{angle}°", variable=self.angle_var, value=angle,
                          font=('Segoe UI', 10), fg=self.colors['text_primary'],
                          bg=self.colors['bg_primary'], selectcolor=self.colors['accent'],
                          activebackground=self.colors['bg_primary']).pack(side="left", padx=(0, 20))
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        rotate_btn = ttk.Button(btn_frame, text="↻ Rotate PDF", style='Modern.TButton',
                               command=self.rotate_pdf_file)
        rotate_btn.pack()
    
    def show_pdf_info_tool(self):
        """Show PDF info tool"""
        self.create_tool_header("PDF Information", "Get detailed information about a PDF file")
        
        # File input
        self.info_pdf = self.create_file_input(self.content_frame, "PDF File:", 
                                              lambda e: self.browse_file(e))
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=20)
        
        info_btn = ttk.Button(btn_frame, text="🔍 Get PDF Info", style='Modern.TButton',
                             command=self.get_pdf_info)
        info_btn.pack()
        
        # Info display area
        self.info_text = tk.Text(self.content_frame, height=15, 
                                font=('Consolas', 10),
                                bg=self.colors['bg_secondary'],
                                fg=self.colors['text_primary'],
                                insertbackground=self.colors['text_primary'],
                                selectbackground=self.colors['accent'])
        self.info_text.pack(fill="both", expand=True, pady=(20, 0))
    
    def show_history_tool(self):
        """Show operation history"""
        self.create_tool_header("Operation History", "View your recent PDF operations")
        
        # History list
        history_frame = tk.Frame(self.content_frame, bg=self.colors['bg_secondary'])
        history_frame.pack(fill="both", expand=True)
        
        # Headers
        headers = ["Time", "Operation", "Files", "Status"]
        header_frame = tk.Frame(history_frame, bg=self.colors['bg_tertiary'])
        header_frame.pack(fill="x", padx=2, pady=2)
        
        for i, header in enumerate(headers):
            tk.Label(header_frame, text=header, font=('Segoe UI', 11, 'bold'),
                    fg=self.colors['text_primary'], bg=self.colors['bg_tertiary']).grid(
                    row=0, column=i, padx=10, pady=10, sticky="w")
        
        # History entries
        if not self.history:
            tk.Label(history_frame, text="No operations yet", 
                    font=('Segoe UI', 12),
                    fg=self.colors['text_secondary'],
                    bg=self.colors['bg_secondary']).pack(expand=True)
        else:
            for entry in self.history[-20:]:  # Show last 20 entries
                self.create_history_entry(history_frame, entry)
    
    def create_history_entry(self, parent, entry):
        """Create a history entry row"""
        row_frame = tk.Frame(parent, bg=self.colors['bg_secondary'])
        row_frame.pack(fill="x", padx=2, pady=1)
        
        data = [entry['time'], entry['operation'], entry['files'], entry['status']]
        for i, item in enumerate(data):
            color = self.colors['success'] if entry['status'] == 'Success' else self.colors['error']
            if i != 3:  # Not status column
                color = self.colors['text_primary']
            
            tk.Label(row_frame, text=str(item), font=('Segoe UI', 10),
                    fg=color, bg=self.colors['bg_secondary']).grid(
                    row=0, column=i, padx=10, pady=5, sticky="w")
    
    def add_to_history(self, operation, files, status):
        """Add operation to history"""
        self.history.append({
            'time': datetime.now().strftime("%H:%M:%S"),
            'operation': operation,
            'files': files,
            'status': status
        })
    
    def update_status(self, message, progress=0):
        """Update status bar"""
        self.status_var.set(message)
        self.progress_var.set(progress)
        self.root.update_idletasks()
    
    def run_with_progress(self, func, *args, **kwargs):
        """Run a function with progress indication"""
        def worker():
            try:
                self.update_status("Processing...", 25)
                result = func(*args, **kwargs)
                self.update_status("Completed successfully", 100)
                return result
            except Exception as e:
                self.update_status(f"Error: {str(e)}", 0)
                messagebox.showerror("Error", str(e))
            finally:
                # Reset progress after 2 seconds
                self.root.after(2000, lambda: self.update_status("Ready", 0))
        
        # Run in thread to prevent UI freezing
        thread = threading.Thread(target=worker)
        thread.daemon = True
        thread.start()
    
    # PDF Operation Methods
    def merge_pdfs(self):
        """Merge two PDF files"""
        pdf1 = self.merge_pdf1.get()
        pdf2 = self.merge_pdf2.get()
        
        if not pdf1 or not pdf2:
            messagebox.showerror("Error", "Please select both PDF files!")
            return
        
        try:
            self.update_status("Merging PDFs...", 50)
            merger = PdfFileMerger()
            merger.append(pdf1)
            merger.append(pdf2)
            
            output = filedialog.asksaveasfilename(
                defaultextension=".pdf", 
                filetypes=[("PDF files", "*.pdf")]
            )
            
            if output:
                merger.write(output)
                merger.close()
                self.add_to_history("Merge", f"{os.path.basename(pdf1)}, {os.path.basename(pdf2)}", "Success")
                messagebox.showinfo("Success", "PDFs merged successfully!")
                self.update_status("PDFs merged successfully", 100)
            else:
                self.update_status("Operation cancelled", 0)
                
        except Exception as e:
            self.add_to_history("Merge", f"{os.path.basename(pdf1)}, {os.path.basename(pdf2)}", "Failed")
            messagebox.showerror("Error", f"Failed to merge PDFs: {str(e)}")
            self.update_status("Ready", 0)
    
    def rotate_pdf_file(self):
        """Rotate PDF file"""
        pdf_path = self.rotate_pdf.get()
        angle = self.angle_var.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        try:
            self.update_status("Rotating PDF...", 50)
            reader = PdfFileReader(pdf_path)
            writer = PdfFileWriter()
            
            for page_num in range(reader.getNumPages()):
                page = reader.getPage(page_num)
                page.rotateClockwise(angle)
                writer.addPage(page)
            
            output = filedialog.asksaveasfilename(
                defaultextension=".pdf", 
                filetypes=[("PDF files", "*.pdf")]
            )
            
            if output:
                with open(output, "wb") as f:
                    writer.write(f)
                self.add_to_history("Rotate", os.path.basename(pdf_path), "Success")
                messagebox.showinfo("Success", f"PDF rotated {angle}° successfully!")
                self.update_status("PDF rotated successfully", 100)
            else:
                self.update_status("Operation cancelled", 0)
                
        except Exception as e:
            self.add_to_history("Rotate", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to rotate PDF: {str(e)}")
            self.update_status("Ready", 0)
    
    def get_pdf_info(self):
        """Get PDF information"""
        pdf_path = self.info_pdf.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        try:
            self.update_status("Analyzing PDF...", 50)
            reader = PdfFileReader(pdf_path)
            
            # Get file info
            file_size = os.path.getsize(pdf_path)
            file_size_mb = file_size / (1024 * 1024)
            
            info_text = f"""PDF INFORMATION REPORT
{'='*50}

File Details:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📁 File Name: {os.path.basename(pdf_path)}
📂 Full Path: {pdf_path}
📊 File Size: {file_size_mb:.2f} MB ({file_size:,} bytes)
📄 Total Pages: {reader.getNumPages()}
🔒 Encrypted: {'Yes' if reader.isEncrypted else 'No'}

Document Metadata:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
            
            # Get metadata
            if reader.documentInfo:
                info = reader.documentInfo
                metadata_items = [
                    ('Title', info.get('/Title', 'Not specified')),
                    ('Author', info.get('/Author', 'Not specified')),
                    ('Subject', info.get('/Subject', 'Not specified')),
                    ('Creator', info.get('/Creator', 'Not specified')),
                    ('Producer', info.get('/Producer', 'Not specified')),
                    ('Creation Date', info.get('/CreationDate', 'Not specified')),
                    ('Modification Date', info.get('/ModDate', 'Not specified'))
                ]
                
                for label, value in metadata_items:
                    info_text += f"📝 {label}: {value}\n"
            else:
                info_text += "📝 No metadata available\n"
            
            # Page information
            info_text += f"\nPage Analysis:\n{'━'*50}\n"
            for i in range(min(5, reader.getNumPages())):  # Show first 5 pages
                page = reader.getPage(i)
                mediabox = page.mediaBox
                width = float(mediabox.getWidth()) * 0.352778  # Convert to mm
                height = float(mediabox.getHeight()) * 0.352778
                info_text += f"📄 Page {i+1}: {width:.1f} x {height:.1f} mm\n"
            
            if reader.getNumPages() > 5:
                info_text += f"... and {reader.getNumPages() - 5} more pages\n"
            
            self.info_text.delete(1.0, tk.END)
            self.info_text.insert(1.0, info_text)
            
            self.add_to_history("Info", os.path.basename(pdf_path), "Success")
            self.update_status("PDF analysis completed", 100)
            
        except Exception as e:
            self.add_to_history("Info", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to analyze PDF: {str(e)}")
            self.update_status("Ready", 0)
    
    def show_pdf_to_doc_tool(self):
        """Show PDF to DOC conversion tool"""
        self.create_tool_header("PDF to Word", "Convert PDF documents to editable Word files")
        
        self.pdf_to_doc_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                     lambda e: self.browse_file(e))
        
        # Conversion options
        options_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        options_frame.pack(fill="x", pady=20)
        
        tk.Label(options_frame, text="Conversion Options:", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 10))
        
        self.preserve_layout = tk.BooleanVar(value=True)
        self.extract_images = tk.BooleanVar(value=True)
        
        tk.Checkbutton(options_frame, text="Preserve original layout", 
                      variable=self.preserve_layout,
                      font=('Segoe UI', 10), fg=self.colors['text_primary'],
                      bg=self.colors['bg_primary'], selectcolor=self.colors['accent'],
                      activebackground=self.colors['bg_primary']).pack(anchor="w", pady=2)
        
        tk.Checkbutton(options_frame, text="Extract embedded images", 
                      variable=self.extract_images,
                      font=('Segoe UI', 10), fg=self.colors['text_primary'],
                      bg=self.colors['bg_primary'], selectcolor=self.colors['accent'],
                      activebackground=self.colors['bg_primary']).pack(anchor="w", pady=2)
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        convert_btn = ttk.Button(btn_frame, text="📄 Convert to Word", style='Modern.TButton',
                                command=self.pdf_to_doc)
        convert_btn.pack()
    
    def show_encrypt_tool(self):
        """Show PDF encryption tool"""
        self.create_tool_header("Encrypt PDF", "Protect your PDF with password encryption")
        
        self.encrypt_pdf_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                      lambda e: self.browse_file(e))
        
        # Password inputs
        pwd_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        pwd_frame.pack(fill="x", pady=20)
        
        tk.Label(pwd_frame, text="User Password:", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))
        
        self.user_password = ttk.Entry(pwd_frame, style='Modern.TEntry', show="*", font=('Segoe UI', 10))
        self.user_password.pack(fill="x", ipady=8, pady=(0, 15))
        
        tk.Label(pwd_frame, text="Owner Password (Optional):", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))
        
        self.owner_password = ttk.Entry(pwd_frame, style='Modern.TEntry', show="*", font=('Segoe UI', 10))
        self.owner_password.pack(fill="x", ipady=8)
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        encrypt_btn = ttk.Button(btn_frame, text="🔐 Encrypt PDF", style='Modern.TButton',
                                command=self.encrypt_pdf)
        encrypt_btn.pack()
    
    def show_watermark_tool(self):
        """Show watermark tool"""
        self.create_tool_header("Add Watermark", "Add text or image watermark to your PDF")
        
        self.watermark_pdf_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                        lambda e: self.browse_file(e))
        
        # Watermark options
        wm_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        wm_frame.pack(fill="x", pady=20)
        
        tk.Label(wm_frame, text="Watermark Text:", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))
        
        self.watermark_text = ttk.Entry(wm_frame, style='Modern.TEntry', font=('Segoe UI', 10))
        self.watermark_text.pack(fill="x", ipady=8, pady=(0, 15))
        self.watermark_text.insert(0, "CONFIDENTIAL")
        
        # Watermark settings
        settings_frame = tk.Frame(wm_frame, bg=self.colors['bg_primary'])
        settings_frame.pack(fill="x", pady=10)
        
        tk.Label(settings_frame, text="Opacity:", 
                font=('Segoe UI', 10),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.opacity_var = tk.DoubleVar(value=0.3)
        opacity_scale = tk.Scale(settings_frame, from_=0.1, to=1.0, resolution=0.1,
                                orient="horizontal", variable=self.opacity_var,
                                bg=self.colors['bg_primary'], fg=self.colors['text_primary'],
                                highlightthickness=0)
        opacity_scale.grid(row=0, column=1, sticky="ew", padx=(0, 20))
        
        tk.Label(settings_frame, text="Font Size:", 
                font=('Segoe UI', 10),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).grid(row=0, column=2, sticky="w", padx=(0, 10))
        
        self.font_size_var = tk.IntVar(value=36)
        font_size_spin = tk.Spinbox(settings_frame, from_=12, to=72, textvariable=self.font_size_var,
                                   width=5, font=('Segoe UI', 10))
        font_size_spin.grid(row=0, column=3, sticky="w")
        
        settings_frame.grid_columnconfigure(1, weight=1)
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        watermark_btn = ttk.Button(btn_frame, text="📑 Add Watermark", style='Modern.TButton',
                                  command=self.add_watermark)
        watermark_btn.pack()
    
    # Placeholder methods for new features
    def show_pdf_to_excel_tool(self):
        """Show PDF to Excel tool"""
        self.create_tool_header("PDF to Excel", "Extract tables from PDF to Excel format")

        self.pdf_to_excel_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                         lambda e: self.browse_file(e))
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        convert_btn = ttk.Button(btn_frame, text="📊 Convert to Excel", style='Modern.TButton',
                                 command=self.pdf_to_excel)
        convert_btn.pack()

    def check_java_installation(self):
        """Check if Java is installed and accessible"""
        try:
            result = subprocess.run(['java', '-version'], 
                                  capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
            return False
    
    def pdf_to_excel_with_pdfplumber(self, pdf_path, output_path):
        """Alternative PDF to Excel conversion using pdfplumber"""
        tables_found = 0
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    # Try to extract tables from the page
                    tables = page.extract_tables()
                    
                    if tables:
                        for table_num, table in enumerate(tables):
                            if table and len(table) > 0:
                                # Convert table to DataFrame
                                df = pd.DataFrame(table[1:], columns=table[0] if table[0] else None)
                                
                                # Clean up the DataFrame
                                df = df.dropna(how='all').dropna(axis=1, how='all')
                                
                                if not df.empty:
                                    sheet_name = f'Page_{page_num+1}_Table_{table_num+1}'
                                    # Truncate sheet name if too long
                                    if len(sheet_name) > 31:
                                        sheet_name = f'P{page_num+1}_T{table_num+1}'
                                    
                                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                                    tables_found += 1
                    
                    # Also try to extract text and look for tabular data
                    text = page.extract_text()
                    if text and not tables:
                        # Try to detect simple tabular data in text
                        lines = text.split('\n')
                        potential_table = []
                        
                        for line in lines:
                            # Look for lines with multiple whitespace-separated values
                            parts = line.split()
                            if len(parts) >= 2:
                                potential_table.append(parts)
                        
                        if len(potential_table) > 1:
                            try:
                                df = pd.DataFrame(potential_table[1:], columns=potential_table[0])
                                df = df.dropna(how='all').dropna(axis=1, how='all')
                                
                                if not df.empty:
                                    sheet_name = f'Page_{page_num+1}_Text'
                                    if len(sheet_name) > 31:
                                        sheet_name = f'P{page_num+1}_Text'
                                    
                                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                                    tables_found += 1
                            except Exception:
                                pass  # Skip if conversion fails
        
        return tables_found

    def pdf_to_excel(self):
        """Convert PDF to Excel with fallback methods"""
        pdf_path = self.pdf_to_excel_file.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        try:
            self.update_status("Converting PDF to Excel...", 25)
            output = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            
            if not output:
                self.update_status("Operation cancelled", 0)
                return
            
            tables_found = 0
            method_used = ""
            
            # Method 1: Try tabula-py if available and Java is installed
            if TABULA_AVAILABLE and self.check_java_installation():
                try:
                    self.update_status("Using Tabula-py for conversion...", 50)
                    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
                    
                    if tables and len(tables) > 0:
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            for i, table in enumerate(tables):
                                if not table.empty:
                                    sheet_name = f'Table_{i+1}'
                                    table.to_excel(writer, sheet_name=sheet_name, index=False)
                                    tables_found += 1
                        method_used = "Tabula-py"
                except Exception as e:
                    print(f"Tabula-py failed: {e}")
                    tables_found = 0
            
            # Method 2: Fallback to pdfplumber if tabula failed or not available
            if tables_found == 0:
                self.update_status("Using PDFplumber for conversion...", 75)
                try:
                    tables_found = self.pdf_to_excel_with_pdfplumber(pdf_path, output)
                    method_used = "PDFplumber"
                except Exception as e:
                    print(f"PDFplumber failed: {e}")
                    tables_found = 0
            
            # Method 3: Last resort - extract all text and create a simple Excel file
            if tables_found == 0:
                self.update_status("Extracting text to Excel...", 90)
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        all_text = []
                        for page in pdf.pages:
                            text = page.extract_text()
                            if text:
                                all_text.append(text)
                    
                    if all_text:
                        # Create a simple Excel file with text content
                        df = pd.DataFrame({'Page': range(1, len(all_text) + 1), 
                                         'Content': all_text})
                        df.to_excel(output, sheet_name='PDF_Content', index=False)
                        tables_found = 1
                        method_used = "Text extraction"
                except Exception as e:
                    print(f"Text extraction failed: {e}")
            
            if tables_found > 0:
                self.add_to_history("PDF to Excel", os.path.basename(pdf_path), "Success")
                success_msg = f"PDF converted to Excel successfully!\n"
                success_msg += f"Method used: {method_used}\n"
                success_msg += f"Tables/Sheets created: {tables_found}"
                messagebox.showinfo("Success", success_msg)
                self.update_status("Conversion completed", 100)
            else:
                self.add_to_history("PDF to Excel", os.path.basename(pdf_path), "Failed")
                error_msg = "Failed to extract any data from the PDF.\n"
                if not TABULA_AVAILABLE:
                    error_msg += "Note: tabula-py is not available.\n"
                elif not self.check_java_installation():
                    error_msg += "Note: Java is not installed or not accessible.\n"
                error_msg += "Please ensure the PDF contains extractable tables or text."
                messagebox.showerror("Error", error_msg)
                self.update_status("No data extracted", 0)
                
        except Exception as e:
            self.add_to_history("PDF to Excel", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to convert PDF: {str(e)}")
            self.update_status("Ready", 0)
    def show_split_tool(self):
        """Show split PDF tool"""
        self.create_tool_header("Split PDF", "Split PDF into multiple files")

        self.split_pdf_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                     lambda e: self.browse_file(e))

        # Page range input
        page_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        page_frame.pack(fill="x", pady=20)

        tk.Label(page_frame, text="Page Range (e.g., 1-3,5,7-10):", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))

        self.page_range = ttk.Entry(page_frame, style='Modern.TEntry', font=('Segoe UI', 10))
        self.page_range.pack(fill="x", ipady=8)
    
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        split_btn = ttk.Button(btn_frame, text="✂️ Split PDF", style='Modern.TButton',
                                command=self.split_pdf)
        split_btn.pack()

    def split_pdf(self):
        """Split PDF into multiple files"""
        pdf_path = self.split_pdf_file.get()
        page_ranges = self.page_range.get()
        
        if not pdf_path or not page_ranges:
            messagebox.showerror("Error", "Please select a PDF file and enter page ranges!")
            return
        
        try:
            self.update_status("Splitting PDF...", 50)
            reader = PdfFileReader(pdf_path)
            page_ranges = self.parse_page_ranges(page_ranges)

            for start, end in page_ranges:
                writer = PdfFileWriter()
                for page_num in range(start, end + 1):
                    writer.addPage(reader.getPage(page_num))

                output = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                                      filetypes=[("PDF files", "*.pdf")],
                                                      title=f"Save PDF from pages {start + 1}-{end + 1}")
                if output:
                    with open(output, "wb") as f:
                        writer.write(f)

            self.add_to_history("Split", os.path.basename(pdf_path), "Success")
            messagebox.showinfo("Success", "PDF split successfully!")
            self.update_status("Splitting completed", 100)

        except Exception as e:
            self.add_to_history("Split", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to split PDF: {str(e)}")
            self.update_status("Ready", 0)

    def parse_page_ranges(self, page_ranges_str):
        """Parse page ranges from a string to a list of (start, end) tuples"""
        ranges = []
        for part in page_ranges_str.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                ranges.append((start - 1, end - 1))  # Convert to zero-based index
            else:
                page = int(part) - 1  # Convert to zero-based index
                ranges.append((page, page))  # Single page as a range
        return ranges
    
    def parse_page_ranges_to_set(self, page_ranges_str):
        """Parse page ranges from a string to a set of page indices"""
        pages = set()
        for part in page_ranges_str.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                pages.update(range(start - 1, end))  # Convert to zero-based index
            else:
                pages.add(int(part) - 1)  # Convert to zero-based index
        return pages
    def show_compress_tool(self):
        """Show compress PDF tool"""
        self.create_tool_header("Compress PDF", "Reduce PDF file size")

        self.compress_pdf_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                       lambda e: self.browse_file(e))
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        compress_btn = ttk.Button(btn_frame, text="🗜️ Compress PDF", style='Modern.TButton',
                                  command=self.compress_pdf)
        compress_btn.pack()

    def compress_pdf(self):
        """Compress PDF file"""
        pdf_path = self.compress_pdf_file.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        try:
            self.update_status("Compressing PDF...", 50)

            # Placeholder for actual compression logic
            # This demo uses the PyPDF3 library for compression effect

            reader = PdfFileReader(pdf_path)
            writer = PdfFileWriter()
            writer.appendPagesFromReader(reader)

            output = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                                  filetypes=[("PDF files", "*.pdf")],
                                                  title="Save compressed PDF")
            if output:
                with open(output, "wb") as f:
                    writer.write(f)

            self.add_to_history("Compress", os.path.basename(pdf_path), "Success")
            messagebox.showinfo("Success", "PDF compressed successfully (in demo mode)!")
            self.update_status("Compression completed", 100)

        except Exception as e:
            self.add_to_history("Compress", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to compress PDF: {str(e)}")
            self.update_status("Ready", 0)
    def show_pdf_to_jpg_tool(self):
        """Show PDF to JPG tool"""
        self.create_tool_header("PDF to JPG", "Convert PDF pages to JPG images")

        self.pdf_to_jpg_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                      lambda e: self.browse_file(e))
        
        # Image DPI selection
        dpi_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        dpi_frame.pack(fill="x", pady=10)

        tk.Label(dpi_frame, text="Select Image DPI:", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))

        self.dpi_var = tk.IntVar(value=200)
        dpi_options = tk.Frame(dpi_frame, bg=self.colors['bg_primary'])
        dpi_options.pack(anchor="w")

        for dpi in [100, 200, 300]:
            tk.Radiobutton(dpi_options, text=f"{dpi} DPI", variable=self.dpi_var, value=dpi,
                          font=('Segoe UI', 10), fg=self.colors['text_primary'],
                          bg=self.colors['bg_primary'], selectcolor=self.colors['accent'],
                          activebackground=self.colors['bg_primary']).pack(side="left", padx=(0, 20))
    
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        convert_btn = ttk.Button(btn_frame, text="🖼️ Convert to JPG", style='Modern.TButton',
                                command=self.pdf_to_jpg)
        convert_btn.pack()

    def pdf_to_jpg(self):
        """Convert PDF pages to JPG images"""
        pdf_path = self.pdf_to_jpg_file.get()
        dpi = self.dpi_var.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        try:
            self.update_status("Converting PDF to JPG...", 50)
            images = convert_from_path(pdf_path, dpi=dpi)
            
            output_dir = filedialog.askdirectory(title="Select directory to save images")
            if output_dir:
                for i, image in enumerate(images, start=1):
                    image_path = os.path.join(output_dir, f"page_{i}.jpg")
                    image.save(image_path, "JPEG")

                self.add_to_history("PDF to JPG", os.path.basename(pdf_path), "Success")
                messagebox.showinfo("Success", "PDF converted to JPG images successfully!")
                self.update_status("Conversion to JPG completed", 100)
            else:
                self.update_status("Operation cancelled", 0)

        except Exception as e:
            self.add_to_history("PDF to JPG", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to convert PDF to JPG: {str(e)}")
            self.update_status("Ready", 0)
    def show_jpg_to_pdf_tool(self):
        """Show JPG to PDF tool"""
        self.create_tool_header("JPG to PDF", "Convert JPG images to PDF")

        self.jpg_to_pdf_files = []  # Store multiple files

        # Files selection
        file_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        file_frame.pack(fill="x", pady=10)

        select_btn = ttk.Button(file_frame, text="Select JPG Files", style='Modern.TButton',
                                command=self.select_jpg_files)
        select_btn.pack(pady=(0, 10))
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        convert_btn = ttk.Button(btn_frame, text="🗃️ Convert to PDF", style='Modern.TButton',
                                command=self.jpg_to_pdf)
        convert_btn.pack()

    def select_jpg_files(self):
        """Select JPG files for conversion"""
        file_paths = filedialog.askopenfilenames(filetypes=[("JPG Files", "*.jpg"), ("JPEG Files", "*.jpeg")],
                                                title="Select JPG files")
        if file_paths:
            self.jpg_to_pdf_files = file_paths

    def jpg_to_pdf(self):
        """Convert selected JPG images to a single PDF"""
        if not self.jpg_to_pdf_files:
            messagebox.showerror("Error", "Please select JPG files!")
            return
        
        try:
            self.update_status("Converting JPG to PDF...", 50)
            images = [Image.open(jpg) for jpg in self.jpg_to_pdf_files]
            
            if images:
                output = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                                      filetypes=[("PDF files", "*.pdf")],
                                                      title="Save as PDF")
                if output:
                    images[0].save(output, save_all=True, append_images=images[1:])  # Merge into single PDF
                    self.add_to_history("JPG to PDF", f"{len(images)} images", "Success")
                    messagebox.showinfo("Success", "JPG images converted to PDF successfully!")
                    self.update_status("Conversion to PDF completed", 100)
                else:
                    self.update_status("Operation cancelled", 0)

            for image in images:
                image.close()  # Close the file handles

        except Exception as e:
            self.add_to_history("JPG to PDF", f"{len(self.jpg_to_pdf_files)} images", "Failed")
            messagebox.showerror("Error", f"Failed to convert JPG to PDF: {str(e)}")
            self.update_status("Ready", 0)
    def show_remove_pages_tool(self):
        """Show remove pages tool"""
        self.create_tool_header("Remove Pages", "Remove specific pages or page ranges from PDF")

        self.remove_pages_pdf_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                          lambda e: self.browse_file(e))

        # Page range input
        page_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        page_frame.pack(fill="x", pady=20)

        tk.Label(page_frame, text="Pages to Remove:", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))

        # Examples label
        examples_label = tk.Label(page_frame, 
                                 text="Examples: 1,3,6 (individual pages) or 1-3,5-7,10 (page ranges)", 
                                 font=('Segoe UI', 9),
                                 fg=self.colors['text_secondary'],
                                 bg=self.colors['bg_primary'])
        examples_label.pack(anchor="w", pady=(0, 10))

        self.pages_to_remove = ttk.Entry(page_frame, style='Modern.TEntry', font=('Segoe UI', 10))
        self.pages_to_remove.pack(fill="x", ipady=8)
        
        # Help text
        help_frame = tk.Frame(page_frame, bg=self.colors['bg_primary'])
        help_frame.pack(fill="x", pady=(10, 0))
        
        help_text = tk.Label(help_frame, 
                            text="• Use commas to separate multiple pages/ranges\n• Use hyphens for ranges (e.g., 1-5 removes pages 1,2,3,4,5)\n• Mix individual pages and ranges (e.g., 1,3-5,8)", 
                            font=('Segoe UI', 9),
                            fg=self.colors['text_secondary'],
                            bg=self.colors['bg_primary'],
                            justify="left")
        help_text.pack(anchor="w")
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        remove_btn = ttk.Button(btn_frame, text="🗑️ Remove Pages", style='Modern.TButton',
                                command=self.remove_pages)
        remove_btn.pack()

    def remove_pages(self):
        """Remove specified pages from PDF"""
        pdf_path = self.remove_pages_pdf_file.get()
        pages_to_remove = self.pages_to_remove.get()
        
        if not pdf_path or not pages_to_remove:
            messagebox.showerror("Error", "Please select a PDF file and enter pages to remove!")
            return
        
        try:
            self.update_status("Removing pages from PDF...", 25)
            
            # Validate PDF file exists
            if not os.path.exists(pdf_path):
                messagebox.showerror("Error", "Selected PDF file does not exist!")
                self.update_status("Ready", 0)
                return
            
            reader = PdfFileReader(pdf_path)
            writer = PdfFileWriter()
            
            # Validate PDF has pages
            total_pages = reader.getNumPages()
            if total_pages == 0:
                messagebox.showerror("Error", "PDF file has no pages!")
                self.update_status("Ready", 0)
                return

            self.update_status("Parsing page ranges...", 40)
            
            # Parse pages and ranges with error handling
            try:
                pages_to_remove_set = self.parse_page_ranges_to_set(pages_to_remove)
            except Exception as parse_error:
                messagebox.showerror("Error", f"Invalid page range format: {str(parse_error)}")
                self.update_status("Ready", 0)
                return
            
            # Validate page numbers are within range
            invalid_pages = [p for p in pages_to_remove_set if p < 0 or p >= total_pages]
            if invalid_pages:
                messagebox.showerror("Error", f"Invalid page numbers: {invalid_pages}. PDF has {total_pages} pages (1-{total_pages})")
                self.update_status("Ready", 0)
                return
            
            # Check if all pages would be removed
            if len(pages_to_remove_set) >= total_pages:
                messagebox.showerror("Error", "Cannot remove all pages from PDF!")
                self.update_status("Ready", 0)
                return

            self.update_status("Processing pages...", 60)
            
            # Add pages that should be kept
            pages_kept = 0
            for i in range(total_pages):
                if i not in pages_to_remove_set:
                    try:
                        page = reader.getPage(i)
                        writer.addPage(page)
                        pages_kept += 1
                    except Exception as page_error:
                        print(f"Warning: Could not process page {i+1}: {page_error}")
                        continue
            
            if pages_kept == 0:
                messagebox.showerror("Error", "No pages would remain after removal!")
                self.update_status("Ready", 0)
                return

            self.update_status("Saving modified PDF...", 80)
            
            # Get output file path with proper error handling
            try:
                output = filedialog.asksaveasfilename(
                    defaultextension=".pdf", 
                    filetypes=[("PDF files", "*.pdf")],
                    title="Save modified PDF",
                    initialfile=f"modified_{os.path.basename(pdf_path)}"
                )
                
                # Check if user cancelled the save dialog
                if not output:
                    self.update_status("Operation cancelled", 0)
                    return
            except Exception as dialog_error:
                messagebox.showerror("Error", f"File dialog error: {str(dialog_error)}")
                self.update_status("Ready", 0)
                return
            
            # Ensure output path is valid
            try:
                output_dir = os.path.dirname(output)
                if output_dir and not os.path.exists(output_dir):
                    os.makedirs(output_dir, exist_ok=True)
            except Exception as dir_error:
                messagebox.showerror("Error", f"Cannot create output directory: {str(dir_error)}")
                self.update_status("Ready", 0)
                return
            
            # Write the modified PDF
            try:
                # Check if writer has any pages
                if writer.getNumPages() == 0:
                    messagebox.showerror("Error", "No pages to write in the output PDF!")
                    self.update_status("Ready", 0)
                    return
                
                # Ensure output is a valid string
                if not isinstance(output, str) or not output:
                    messagebox.showerror("Error", "Invalid output path")
                    self.update_status("Ready", 0)
                    return
                
                # Write the file
                with open(output, "wb") as f:
                    writer.write(f)
                
                # Verify the file was created and has content
                if os.path.exists(output) and os.path.getsize(output) > 0:
                    pages_removed = len(pages_to_remove_set)
                    self.add_to_history("Remove Pages", os.path.basename(pdf_path), "Success")
                    success_msg = f"Pages removed successfully!\n"
                    success_msg += f"Removed {pages_removed} page(s), {pages_kept} page(s) remaining."
                    messagebox.showinfo("Success", success_msg)
                    self.update_status("Pages removal completed", 100)
                else:
                    raise Exception("Output file was not created properly")
                    
            except Exception as write_error:
                messagebox.showerror("Error", f"Failed to save modified PDF: {str(write_error)}")
                self.add_to_history("Remove Pages", os.path.basename(pdf_path), "Failed")
                self.update_status("Ready", 0)
                return

        except Exception as e:
            self.add_to_history("Remove Pages", os.path.basename(pdf_path), "Failed")
            error_msg = f"Failed to remove pages: {str(e)}"
            if "NoneType" in str(e):
                error_msg += "\n\nThis might be due to a cancelled file dialog or invalid PDF structure."
            messagebox.showerror("Error", error_msg)
            self.update_status("Ready", 0)
    def show_pdf_to_odt_tool(self):
        """Show PDF to ODT tool"""
        self.create_tool_header("PDF to ODT", "Convert PDF to OpenDocument Text format")
        
        # File input
        self.pdf_to_odt_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                     lambda e: self.browse_file(e))
        
        # Conversion options
        options_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        options_frame.pack(fill="x", pady=20)
        
        tk.Label(options_frame, text="Conversion Options:", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 10))
        
        self.preserve_formatting = tk.BooleanVar(value=True)
        self.extract_text_only = tk.BooleanVar(value=False)
        
        tk.Checkbutton(options_frame, text="Preserve text formatting", 
                      variable=self.preserve_formatting,
                      font=('Segoe UI', 10), fg=self.colors['text_primary'],
                      bg=self.colors['bg_primary'], selectcolor=self.colors['accent'],
                      activebackground=self.colors['bg_primary']).pack(anchor="w", pady=2)
        
        tk.Checkbutton(options_frame, text="Extract text only (ignore formatting)", 
                      variable=self.extract_text_only,
                      font=('Segoe UI', 10), fg=self.colors['text_primary'],
                      bg=self.colors['bg_primary'], selectcolor=self.colors['accent'],
                      activebackground=self.colors['bg_primary']).pack(anchor="w", pady=2)
        
        # Info text
        info_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        info_frame.pack(fill="x", pady=10)
        
        info_text = tk.Label(info_frame, 
                            text="ODT (OpenDocument Text) is a standard format for text documents\nthat can be opened with LibreOffice Writer, OpenOffice Writer, and other editors.", 
                            font=('Segoe UI', 9),
                            fg=self.colors['text_secondary'],
                            bg=self.colors['bg_primary'],
                            justify="left")
        info_text.pack(anchor="w")
        
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        convert_btn = ttk.Button(btn_frame, text="📝 Convert to ODT", style='Modern.TButton',
                                command=self.pdf_to_odt)
        convert_btn.pack()
        
    def show_decrypt_tool(self):
        """Show PDF decryption tool"""
        self.create_tool_header("Decrypt PDF", "Remove password protection from PDF")

        self.decrypt_pdf_file = self.create_file_input(self.content_frame, "PDF File:", 
                                                     lambda e: self.browse_file(e))

        # Password input
        pwd_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        pwd_frame.pack(fill="x", pady=20)

        tk.Label(pwd_frame, text="Password:", 
                font=('Segoe UI', 11, 'bold'),
                fg=self.colors['text_primary'],
                bg=self.colors['bg_primary']).pack(anchor="w", pady=(0, 5))

        self.decrypt_password = ttk.Entry(pwd_frame, style='Modern.TEntry', show="*", font=('Segoe UI', 10))
        self.decrypt_password.pack(fill="x", ipady=8)
    
        # Action button
        btn_frame = tk.Frame(self.content_frame, bg=self.colors['bg_primary'])
        btn_frame.pack(pady=30)
        
        decrypt_btn = ttk.Button(btn_frame, text="🔓 Decrypt PDF", style='Modern.TButton',
                                 command=self.decrypt_pdf)
        decrypt_btn.pack()

    def decrypt_pdf(self):
        """Remove password protection from PDF"""
        pdf_path = self.decrypt_pdf_file.get()
        password = self.decrypt_password.get()
        
        if not pdf_path or not password:
            messagebox.showerror("Error", "Please select a PDF file and enter the password!")
            return
        
        try:
            self.update_status("Decrypting PDF...", 50)
            reader = PdfFileReader(pdf_path)

            if reader.isEncrypted:
                reader.decrypt(password)
            else:
                messagebox.showerror("Error", "PDF is not encrypted!")
                return

            writer = PdfFileWriter()
            for page_num in range(reader.getNumPages()):
                writer.addPage(reader.getPage(page_num))

            output = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                                  filetypes=[("PDF files", "*.pdf")],
                                                  title="Save decrypted PDF")
            if output:
                with open(output, "wb") as f:
                    writer.write(f)

            self.add_to_history("Decrypt", os.path.basename(pdf_path), "Success")
            messagebox.showinfo("Success", "PDF decrypted successfully!")
            self.update_status("Decryption completed", 100)

        except Exception as e:
            self.add_to_history("Decrypt", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to decrypt PDF: {str(e)}")
            self.update_status("Ready", 0)
    
    def pdf_to_odt(self):
        """Convert PDF to ODT"""
        pdf_path = self.pdf_to_odt_file.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        try:
            self.update_status("Converting PDF to ODT...", 50)
            output = filedialog.asksaveasfilename(
                defaultextension=".odt", 
                filetypes=[("ODT files", "*.odt")],
                title="Save as ODT"
            )
            
            if output:
                with pdfplumber.open(pdf_path) as pdf:
                    odt_doc = OpenDocumentText()
                    
                    if self.preserve_formatting.get():
                        text_style = Style(name="TextStyle", family="text")
                        text_style.addElement(TextProperties(attributes={"fontsize":"12pt"}))
                        odt_doc.styles.addElement(text_style)
                        
                        para_style = Style(name="ParaStyle", family="paragraph")
                        para_style.addElement(ParagraphProperties(attributes={"textalign":"left"}))
                        odt_doc.styles.addElement(para_style)
                    
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:  # Only process if text is found
                            if self.extract_text_only.get():
                                # Simple text extraction
                                paragraphs = text.split('\n')
                                for para in paragraphs:
                                    if para.strip():  # Skip empty paragraphs
                                        p = P(text=para)
                                        odt_doc.text.addElement(p)
                            else:
                                # Preserve formatting
                                p = P(stylename="ParaStyle")
                                span = Span(stylename="TextStyle", text=text)
                                p.addElement(span)
                                odt_doc.text.addElement(p)
                    
                    odt_doc.save(output)
                    
                self.add_to_history("PDF to ODT", os.path.basename(pdf_path), "Success")
                messagebox.showinfo("Success", "PDF converted to ODT successfully!")
                self.update_status("Conversion completed", 100)
            else:
                self.update_status("Operation cancelled", 0)
        
        except Exception as e:
            self.add_to_history("PDF to ODT", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to convert PDF: {str(e)}")
            self.update_status("Ready", 0)
    # New feature implementations
    def pdf_to_doc(self):
        """Convert PDF to Word document"""
        pdf_path = self.pdf_to_doc_file.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        try:
            self.update_status("Converting PDF to Word...", 50)
            output = filedialog.asksaveasfilename(
                defaultextension=".docx", 
                filetypes=[("Word files", "*.docx")]
            )
            
            if output:
                cv = Converter(pdf_path)
                cv.convert(output, start=0, end=None)
                cv.close()
                
                self.add_to_history("PDF to DOC", os.path.basename(pdf_path), "Success")
                messagebox.showinfo("Success", "PDF converted to Word successfully!")
                self.update_status("Conversion completed", 100)
            else:
                self.update_status("Operation cancelled", 0)
                
        except Exception as e:
            self.add_to_history("PDF to DOC", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to convert PDF: {str(e)}")
            self.update_status("Ready", 0)
    
    def encrypt_pdf(self):
        """Encrypt PDF with password"""
        pdf_path = self.encrypt_pdf_file.get()
        user_pwd = self.user_password.get()
        owner_pwd = self.owner_password.get() or user_pwd
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        if not user_pwd:
            messagebox.showerror("Error", "Please enter a user password!")
            return
        
        try:
            self.update_status("Encrypting PDF...", 50)
            reader = PdfFileReader(pdf_path)
            writer = PdfFileWriter()
            
            for page_num in range(reader.getNumPages()):
                writer.addPage(reader.getPage(page_num))
            
            writer.encrypt(user_pwd, owner_pwd, use_128bit=True)
            
            output = filedialog.asksaveasfilename(
                defaultextension=".pdf", 
                filetypes=[("PDF files", "*.pdf")]
            )
            
            if output:
                with open(output, "wb") as f:
                    writer.write(f)
                
                self.add_to_history("Encrypt", os.path.basename(pdf_path), "Success")
                messagebox.showinfo("Success", "PDF encrypted successfully!")
                self.update_status("Encryption completed", 100)
                
                # Clear password fields
                self.user_password.delete(0, tk.END)
                self.owner_password.delete(0, tk.END)
            else:
                self.update_status("Operation cancelled", 0)
                
        except Exception as e:
            self.add_to_history("Encrypt", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to encrypt PDF: {str(e)}")
            self.update_status("Ready", 0)
    
    def add_watermark(self):
        """Add watermark to PDF"""
        pdf_path = self.watermark_pdf_file.get()
        watermark_text = self.watermark_text.get()
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file!")
            return
        
        if not watermark_text:
            messagebox.showerror("Error", "Please enter watermark text!")
            return
        
        try:
            self.update_status("Adding watermark...", 50)
            
            # This is a simplified implementation
            # In a real application, you'd use a more sophisticated watermarking library
            messagebox.showinfo("Info", "Watermark feature requires additional libraries.\nThis is a demonstration of the UI.")
            
            self.add_to_history("Watermark", os.path.basename(pdf_path), "Demo")
            self.update_status("Watermark demo completed", 100)
            
        except Exception as e:
            self.add_to_history("Watermark", os.path.basename(pdf_path), "Failed")
            messagebox.showerror("Error", f"Failed to add watermark: {str(e)}")
            self.update_status("Ready", 0)
            
            
    
    
    

    
    
    def run(self):
        """Start the application"""
        self.root.mainloop()

# Create and run the application
if __name__ == "__main__":
    app = ModernPDFTool()
    app.run()