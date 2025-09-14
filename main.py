import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from converters import FileConverter
import threading
import shutil

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal File Converter")
        self.root.geometry("900x700")
        self.root.configure(bg='#2c3e50')
        self.root.resizable(True, True)
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('Title.TLabel', font=('Arial', 20, 'bold'), foreground='#ecf0f1')
        self.style.configure('Heading.TLabel', font=('Arial', 12, 'bold'), foreground='#34495e')
        self.style.configure('Convert.TButton', font=('Arial', 14, 'bold'), padding=15)
        self.style.configure('Download.TButton', font=('Arial', 12, 'bold'), padding=10)
        
        self.converter = FileConverter()
        self.converted_file_path = None
        self.converted_file_data = None
        self.converted_file_name = None
        self.setup_ui()
        
    def setup_ui(self):
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Main container with gradient-like background
        main_container = tk.Frame(self.root, bg='#34495e')
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=0, pady=0)
        main_container.columnconfigure(1, weight=1)
        
        # Header section
        header_frame = tk.Frame(main_container, bg='#2c3e50', height=80)
        header_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.grid_propagate(False)
        
        title_label = tk.Label(header_frame, text="🔄 Universal File Converter", 
                             font=('Arial', 24, 'bold'), fg='#ecf0f1', bg='#2c3e50')
        title_label.pack(expand=True)
        
        # Content container
        content_frame = tk.Frame(main_container, bg='#ecf0f1', relief=tk.RAISED, bd=2)
        content_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), padx=20, pady=(0, 20))
        content_frame.columnconfigure(1, weight=1)
        
        # File selection section
        file_frame = tk.LabelFrame(content_frame, text="📁 File Selection", font=('Arial', 12, 'bold'),
                                 fg='#2c3e50', bg='#ecf0f1', padx=15, pady=15)
        file_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=20, pady=(20, 15))
        file_frame.columnconfigure(1, weight=1)
        
        # Upload area
        upload_frame = tk.Frame(file_frame, bg='#ecf0f1')
        upload_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        upload_frame.columnconfigure(0, weight=1)
        
        self.upload_area = tk.Frame(upload_frame, bg='#ffffff', relief=tk.SOLID, bd=2, height=100)
        self.upload_area.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=10)
        self.upload_area.grid_propagate(False)
        
        upload_label = tk.Label(self.upload_area, text="📤 Drag & Drop or Click to Upload File",
                              font=('Arial', 14, 'bold'), fg='#7f8c8d', bg='#ffffff')
        upload_label.pack(expand=True)
        
        self.file_path = tk.StringVar()
        upload_btn = tk.Button(upload_frame, text="📁 Upload File", command=self.upload_file,
                             font=('Arial', 12, 'bold'), bg='#3498db', fg='white',
                             relief=tk.FLAT, padx=20, pady=8)
        upload_btn.grid(row=1, column=0, pady=10)
        
        # Format selection section
        format_frame = tk.LabelFrame(content_frame, text="🔧 Output Format", font=('Arial', 12, 'bold'),
                                   fg='#2c3e50', bg='#ecf0f1', padx=15, pady=15)
        format_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=20, pady=15)
        
        tk.Label(format_frame, text="Convert to:", font=('Arial', 11, 'bold'),
               fg='#34495e', bg='#ecf0f1').grid(row=0, column=0, sticky=tk.W, pady=10)
        self.output_format = tk.StringVar()
        format_combo = ttk.Combobox(format_frame, textvariable=self.output_format, 
                                  font=('Arial', 11), width=30, state='readonly')
        format_combo['values'] = ('PDF', 'DOCX', 'TXT', 'HTML', 'JPG', 'PNG', 'CSV', 'JSON', 'XML', 'XLSX', 'PPTX', 'PDFA')
        format_combo.grid(row=0, column=1, sticky=tk.W, padx=(15, 0), pady=10)
        
        # Action buttons section
        button_frame = tk.Frame(content_frame, bg='#ecf0f1')
        button_frame.grid(row=2, column=0, columnspan=3, pady=20)
        
        convert_btn = tk.Button(button_frame, text="🚀 Convert File", 
                              command=self.convert_file, font=('Arial', 14, 'bold'),
                              bg='#27ae60', fg='white', relief=tk.FLAT, padx=30, pady=10)
        convert_btn.pack(side=tk.LEFT, padx=10)
        
        self.download_btn = tk.Button(button_frame, text="💾 Download", 
                                    command=self.download_file, font=('Arial', 12, 'bold'),
                                    bg='#e74c3c', fg='white', relief=tk.FLAT, padx=20, pady=8,
                                    state=tk.DISABLED)
        self.download_btn.pack(side=tk.LEFT, padx=10)
        
        # Progress section
        progress_frame = tk.LabelFrame(content_frame, text="📊 Progress", font=('Arial', 12, 'bold'),
                                     fg='#2c3e50', bg='#ecf0f1', padx=15, pady=15)
        progress_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=20, pady=15)
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=500)
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=10)
        
        self.status_label = tk.Label(progress_frame, text="Ready to convert files", 
                                   font=('Arial', 11), fg='#27ae60', bg='#ecf0f1')
        self.status_label.grid(row=1, column=0, pady=5)
        
        # File info section
        info_frame = tk.LabelFrame(content_frame, text="📋 File Information", font=('Arial', 12, 'bold'),
                                 fg='#2c3e50', bg='#ecf0f1', padx=15, pady=15)
        info_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), padx=20, pady=(15, 20))
        info_frame.columnconfigure(0, weight=1)
        info_frame.rowconfigure(0, weight=1)
        
        self.info_text = tk.Text(info_frame, height=10, wrap=tk.WORD, font=('Consolas', 10),
                               bg='#ffffff', fg='#2c3e50', relief=tk.FLAT, bd=2)
        scrollbar = ttk.Scrollbar(info_frame, orient=tk.VERTICAL, command=self.info_text.yview)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        
        self.info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S), pady=5)
        
        self.info_text.insert(tk.END, "🎉 Welcome to Universal File Converter!\n\n")
        self.info_text.insert(tk.END, "📋 Supported formats:\n")
        self.info_text.insert(tk.END, "🖼️ Images: JPG, PNG\n")
        self.info_text.insert(tk.END, "📄 Documents: PDF, DOCX, TXT, HTML, PPTX\n")
        self.info_text.insert(tk.END, "📊 Data: CSV, JSON, XML, XLSX\n")
        self.info_text.insert(tk.END, "🗃️ Archive: PDF/A\n\n")
        self.info_text.insert(tk.END, "📤 Upload a file and choose output format to begin.")
        self.info_text.config(state=tk.DISABLED)
        
    def upload_file(self):
        filetypes = [
            ("All files", "*.*"),
            ("Images", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.webp"),
            ("Documents", "*.pdf *.docx *.txt *.html"),
            ("Data files", "*.csv *.json *.xml *.xlsx"),
            ("Audio files", "*.mp3 *.wav *.ogg *.flac"),
            ("Video files", "*.mp4 *.avi *.mov *.mkv")
        ]
        
        filename = filedialog.askopenfilename(
            title="Upload file to convert",
            filetypes=filetypes
        )
        if filename:
            self.file_path.set(filename)
            self.update_upload_area(filename)
            self.update_file_info(filename)
            
    def update_file_info(self, filepath):
        try:
            file_size = os.path.getsize(filepath) / 1024  # KB
            file_ext = os.path.splitext(filepath)[1].upper()
            
            self.info_text.config(state=tk.NORMAL)
            self.info_text.delete(1.0, tk.END)
            self.info_text.insert(tk.END, f"📤 Uploaded File: {os.path.basename(filepath)}\n")
            self.info_text.insert(tk.END, f"📋 File Type: {file_ext}\n")
            self.info_text.insert(tk.END, f"📏 File Size: {file_size:.1f} KB\n\n")
            self.info_text.insert(tk.END, "✅ Ready for conversion. Select output format and click Convert.")
            self.info_text.config(state=tk.DISABLED)
        except Exception as e:
            print("File info update error:", e)
            
    def update_upload_area(self, filepath):
        try:
            # Clear upload area and show uploaded file
            for widget in self.upload_area.winfo_children():
                widget.destroy()
                
            filename = os.path.basename(filepath)
            file_label = tk.Label(self.upload_area, text=f"📄 {filename}",
                                font=('Arial', 12, 'bold'), fg='#27ae60', bg='#ffffff')
            file_label.pack(expand=True)
            
            # Change upload area color to indicate file is uploaded
            self.upload_area.config(bg='#d5f4e6', relief=tk.SOLID)
        except Exception as e:
            print("Upload area update error:", e)
            
    def convert_file(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Please upload a file to convert")
            return
            
        if not self.output_format.get():
            messagebox.showerror("Error", "Please select an output format")
            return
            
        # Run conversion in separate thread to prevent UI freezing
        thread = threading.Thread(target=self._perform_conversion)
        thread.daemon = True
        thread.start()
        
    def _perform_conversion(self):
        try:
            self.progress.start()
            self.status_label.config(text="Converting...", fg='#f39c12')
            self.download_btn.config(state=tk.DISABLED)
            
            self.info_text.config(state=tk.NORMAL)
            self.info_text.insert(tk.END, "\n\n🔄 Starting conversion...\n")
            self.info_text.config(state=tk.DISABLED)
            self.info_text.see(tk.END)
            
            temp_output_path = self.converter.convert(
                self.file_path.get(), 
                self.output_format.get().lower()
            )
            
            # Read the converted file into memory
            with open(temp_output_path, 'rb') as f:
                self.converted_file_data = f.read()
            
            # Store the filename for download
            self.converted_file_name = os.path.basename(temp_output_path)
            
            # Clean up temporary file
            try:
                os.remove(temp_output_path)
            except:
                pass
            
            self.progress.stop()
            self.status_label.config(text="Conversion completed successfully!", fg='#27ae60')
            
            # Enable download button
            self.download_btn.config(state=tk.NORMAL, bg='#27ae60')
            
            self.info_text.config(state=tk.NORMAL)
            self.info_text.insert(tk.END, f"✅ Conversion completed!\n")
            self.info_text.insert(tk.END, f"📄 File ready for download\n")
            self.info_text.insert(tk.END, f"💾 Click Download to save to Downloads folder\n")
            self.info_text.config(state=tk.DISABLED)
            self.info_text.see(tk.END)
            
            messagebox.showinfo("Success", "File converted successfully!\n\nClick the Download button to save the file.")
            
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="Conversion failed", fg='#e74c3c')
            
            self.info_text.config(state=tk.NORMAL)
            self.info_text.insert(tk.END, f"\n❌ Error: {str(e)}\n")
            self.info_text.config(state=tk.DISABLED)
            self.info_text.see(tk.END)
            
            messagebox.showerror("Error", f"Conversion failed:\n\n{str(e)}")
    
    def download_file(self):
        if not hasattr(self, 'converted_file_data') or not self.converted_file_data:
            messagebox.showerror("Error", "No converted file available. Please convert a file first.")
            return
            
        try:
            # Get Downloads folder path
            downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            if not os.path.exists(downloads_path):
                downloads_path = os.path.expanduser('~')
                
            filename = self.converted_file_name
            download_path = os.path.join(downloads_path, filename)
            
            # Handle duplicate filenames
            counter = 1
            base_name, ext = os.path.splitext(filename)
            while os.path.exists(download_path):
                new_filename = f"{base_name}_{counter}{ext}"
                download_path = os.path.join(downloads_path, new_filename)
                counter += 1
            
            # Write the file data from memory to Downloads
            with open(download_path, 'wb') as f:
                f.write(self.converted_file_data)
            
            # Store the final downloaded file path
            self.converted_file_path = download_path
            
            # Update UI
            self.info_text.config(state=tk.NORMAL)
            self.info_text.insert(tk.END, f"📥 Downloaded to: {download_path}\n")
            self.info_text.config(state=tk.DISABLED)
            self.info_text.see(tk.END)
            
            messagebox.showinfo("Success", f"File downloaded successfully!\n\nSaved to Downloads:\n{os.path.basename(download_path)}")
            
        except Exception as e:
            self.info_text.config(state=tk.NORMAL)
            self.info_text.insert(tk.END, f"❌ Download failed: {str(e)}\n")
            self.info_text.config(state=tk.DISABLED)
            self.info_text.see(tk.END)
            
            messagebox.showerror("Error", f"Download failed:\n\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    
    # Center window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()