import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import os
import docx
import PyPDF2
from bs4 import BeautifulSoup
from tkinter import ttk
import re

class FileSearchApp:
    def __init__(self):
        self.window = ctk.CTk()
        self.window.title("Word Occurrence Counter")
        self.window.geometry("800x600")
        
        # Directory selection frame
        self.dir_frame = ctk.CTkFrame(self.window)
        self.dir_frame.pack(pady=10, padx=10, fill="x")
        
        self.dir_label = ctk.CTkLabel(self.dir_frame, text="Selected Directory:")
        self.dir_label.pack(side="left", padx=5)
        
        self.dir_entry = ctk.CTkEntry(self.dir_frame, width=500)
        self.dir_entry.pack(side="left", padx=5)
        
        self.browse_btn = ctk.CTkButton(self.dir_frame, text="Browse", command=self.browse_directory)
        self.browse_btn.pack(side="left", padx=5)
        
        # Search frame
        self.search_frame = ctk.CTkFrame(self.window)
        self.search_frame.pack(pady=10, padx=10, fill="x")
        
        self.search_label = ctk.CTkLabel(self.search_frame, text="Search Word:")
        self.search_label.pack(side="left", padx=5)
        
        self.search_entry = ctk.CTkEntry(self.search_frame, width=500)
        self.search_entry.pack(side="left", padx=5)
        
        self.search_btn = ctk.CTkButton(self.search_frame, text="Search", command=self.search_files)
        self.search_btn.pack(side="left", padx=5)
        
        # Results frame
        self.results_frame = ctk.CTkFrame(self.window)
        self.results_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        # Create treeview for results
        self.tree = ttk.Treeview(self.results_frame, columns=('File', 'Occurrences'), show='headings')
        self.tree.heading('File', text='File')
        self.tree.heading('Occurrences', text='Occurrences')
        self.tree.column('File', width=600)
        self.tree.column('Occurrences', width=100, anchor='center')
        self.tree.pack(side="left", fill="both", expand=True)
        
        # Add scrollbar to treeview
        self.scrollbar = ttk.Scrollbar(self.results_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        
    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)
            
    def read_txt_file(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as file:
                return file.read()
        except:
            return ""
            
    def read_doc_file(self, filepath):
        try:
            doc = docx.Document(filepath)
            return " ".join([paragraph.text for paragraph in doc.paragraphs])
        except:
            return ""
            
    def read_pdf_file(self, filepath):
        try:
            with open(filepath, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text()
                return text
        except:
            return ""
            
    def read_html_file(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as file:
                soup = BeautifulSoup(file.read(), 'html.parser')
                return soup.get_text()
        except:
            return ""
            
    def search_files(self):
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        search_word = self.search_entry.get().strip()
        directory = self.dir_entry.get()
        
        if not search_word or not directory:
            self.tree.insert('', 'end', values=('Please enter both a search word and select a directory.', ''))
            return
            
        total_occurrences = 0
        files_with_matches = 0
        
        for root, _, files in os.walk(directory):
            for file in files:
                if file.lower().endswith(('.txt', '.doc', '.docx', '.html', '.pdf')):
                    filepath = os.path.join(root, file)
                    content = ""
                    
                    # Read file content based on extension
                    if file.lower().endswith('.txt'):
                        content = self.read_txt_file(filepath)
                    elif file.lower().endswith(('.doc', '.docx')):
                        content = self.read_doc_file(filepath)
                    elif file.lower().endswith('.pdf'):
                        content = self.read_pdf_file(filepath)
                    elif file.lower().endswith('.html'):
                        content = self.read_html_file(filepath)
                        
                    # Count occurrences
                    if content:
                        count = len(re.findall(search_word, content, re.IGNORECASE))
                        if count > 0:
                            # Get relative path if possible
                            try:
                                rel_path = os.path.relpath(filepath, directory)
                            except:
                                rel_path = filepath
                            self.tree.insert('', 'end', values=(rel_path, count))
                            total_occurrences += count
                            files_with_matches += 1
        
        # Insert summary at the top
        if total_occurrences > 0:
            self.tree.insert('', 0, values=('', ''))
            self.tree.insert('', 0, values=(f'Total files with matches: {files_with_matches}', ''))
            self.tree.insert('', 0, values=(f'Total occurrences found: {total_occurrences}', ''))
        else:
            self.tree.insert('', 'end', values=(f'No occurrences of "{search_word}" found in any files.', ''))
            
    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = FileSearchApp()
    app.run()