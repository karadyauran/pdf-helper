from time import sleep
import pdfplumber
import pandas as pd
import re
from typing import List, Dict, Optional
import yaml
import os
import inquirer
from datetime import datetime

class PDFExtractor:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        
    def extract_text(self) -> str:
        text_data = []
        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                text_data.append(page.extract_text())
        return "\n".join(text_data)

class TextParser:
    def __init__(self, title_keywords: List[str]):
        self.title_keywords = title_keywords
        self.title_pattern = re.compile(rf"^({'|'.join(title_keywords)})[\w\s/-]*")
        self.date_pattern = re.compile(r"(Jan|Feb|Mrz|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez) \d{2}( - (Jan|Feb|Mrz|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez) \d{2})?")
    
    def parse_sections(self, text: str) -> Dict[str, List[str]]:
        sections = {}
        current_title = None
        
        for line in text.split("\n"):
            line = line.strip()
            if self.title_pattern.match(line):
                current_title = line
                sections[current_title] = []
            elif current_title and line and not line.startswith("Ernte"):
                sections[current_title].append(line)
        return sections

class DataProcessor:
    def __init__(self, date_pattern):
        self.date_pattern = date_pattern
        
    def process_data(self, sections: Dict[str, List[str]]) -> tuple:
        all_data = []
        max_dates = 0
        max_prices = 0
        
        for title, lines in sections.items():
            for line in lines:
                parts = line.split()
                dates = self.date_pattern.findall(line)
                prices = [p for p in parts if re.search(r"\d{3},\d+", p) and "Prämie" not in p]
                dates = [" ".join(d[:2]).strip() for d in dates]
                
                if dates and prices:
                    max_dates = max(max_dates, len(dates))
                    max_prices = max(max_prices, len(prices))
                    all_data.append([title] + dates + prices)
                    
        return all_data, max_dates, max_prices

class ExcelExporter:
    def __init__(self, output_path: str):
        self.output_path = output_path
        
    def export_to_excel(self, data: List[List[str]], max_dates: int, max_prices: int):
        if not data:
            print("Warnung: Keine Zeilen mit Daten und Preisen gefunden. Bitte überprüfen Sie die PDF-Datei.")
            return
            
        columns = ["Title"] + [f"Date_{i}" for i in range(1, max_dates + 1)] + \
                 [f"Price_{i}" for i in range(1, max_prices + 1)] + ["Average Price"]
        
        for row in data:
            while len(row) < len(columns) - 1:
                row.append(None)
            price_values = [float(p.replace(",", ".")) for p in row[max_dates + 1:] if p is not None]
            row.append(round(sum(price_values) / len(price_values), 2) if price_values else None)
        
        df = pd.DataFrame(data, columns=columns)
        df.to_excel(self.output_path, index=False)

class Settings:
    def __init__(self):
        self.settings_path = "settings.yaml"
        self.default_settings = {
            "title_keywords": ["Weizen", "F-Weizen"]
        }

    def load_settings(self):
        if os.path.exists(self.settings_path):
            with open(self.settings_path, 'r') as f:
                return yaml.safe_load(f)
        else:
            with open(self.settings_path, 'w') as f:
                yaml.dump(self.default_settings, f)
            return self.default_settings

    def save_settings(self, settings):
        with open(self.settings_path, 'w') as f:
            yaml.dump(settings, f)


def clear_console():
    if os.name == 'nt':  # for Windows
        os.system('cls')
    else:  # for Unix/Linux/MacOS
        os.system('clear')

def greet_user():
    print("\nPDF Helper v1.0")
    print("Author: Ocean Script Studio")
    print("-" * 30)
    print()

class MenuHandler:
    def __init__(self):
        self.settings_manager = Settings()
        self.settings = self.settings_manager.load_settings()

    def show_main_menu(self):
        clear_console()
        greet_user()
        questions = [
            inquirer.List('choice',
                 message="Select an option:",
                 choices=[
                     ('Process PDF files', '1'),
                     ('Settings', '2'),
                     ('Exit', '3')
                 ])
        ]
        return inquirer.prompt(questions)['choice']

    def show_settings_menu(self):
        clear_console()
        greet_user()
        print("\nCurrent title keywords:", self.settings["title_keywords"])
        keyword_questions = [
            inquirer.List('action',
                     message="Settings:",
                     choices=[
                     ('Keep current keywords', 'keep'),
                     ('Edit keywords', 'edit')
                     ])
        ]
        return inquirer.prompt(keyword_questions)['action']

    def edit_keywords(self):
        clear_console()
        greet_user()
        new_keywords = []
        while True:
            add_question = [
                inquirer.Text('keyword', message="Enter a keyword (or press Enter to finish)")
            ]
            keyword = inquirer.prompt(add_question)['keyword']
            if not keyword:
                break
            new_keywords.append(keyword)
        return new_keywords

class PDFFileHandler:
    def __init__(self):
        self.pdf_dir = "pdf-files"

    def ensure_pdf_directory(self):
        if not os.path.exists(self.pdf_dir):
            os.makedirs(self.pdf_dir)

    def get_pdf_files(self):
        pdf_files = []
        for file in os.listdir(self.pdf_dir):
            if file.lower().endswith('.pdf'):
                full_path = os.path.join(self.pdf_dir, file)
                pdf_files.append((full_path, os.path.getmtime(full_path)))
        return sorted(pdf_files, key=lambda x: x[1], reverse=True)

    def select_files(self, pdf_files):
        clear_console()
        greet_user()
        choices = [(os.path.basename(f[0]), f[0]) for f in pdf_files]
        questions = [
            inquirer.Checkbox('files',
                         message="Select PDF files to process (use space to select, enter to confirm)",
                         choices=choices)
        ]
        result = inquirer.prompt(questions)
        if result is None:
            return []
        files = result['files']
        if not files:
            return []
        file_data = []
        for pdf_path in files:
            with pdfplumber.open(pdf_path) as pdf:
                file_data.append(pdf_path)
        return file_data

def main():
    menu_handler = MenuHandler()
    pdf_handler = PDFFileHandler()

    while True:
        choice = menu_handler.show_main_menu()
        
        if choice == '1':
            pdf_handler.ensure_pdf_directory()
            pdf_files = pdf_handler.get_pdf_files()

            if not pdf_files:
                print("\nNo PDF files found in the pdf-files directory")
                continue

            print(f"\nFound {len(pdf_files)} PDF files:")
            for file, _ in pdf_files:
                print(f"- {os.path.basename(file)}")

            selected_files = pdf_handler.select_files(pdf_files)

            if 'menu' in selected_files or not selected_files:
                continue

            title_keywords = menu_handler.settings["title_keywords"]
            
            # Create excel-output directory if it doesn't exist
            if not os.path.exists("excel-output"):
                os.makedirs("excel-output")
                
            # Get current date and time for filename
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Process each selected file
            for pdf_path in selected_files:
                pdf_extractor = PDFExtractor(pdf_path)
                pdf_text = pdf_extractor.extract_text()
                
                text_parser = TextParser(title_keywords)
                sections = text_parser.parse_sections(pdf_text)

                data_processor = DataProcessor(text_parser.date_pattern)
                all_data, max_dates, max_prices = data_processor.process_data(sections)

                # Create output filename
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                output_file = f"excel-output/{base_name}_{current_time}.xlsx"
                
                excel_exporter = ExcelExporter(output_file)
                excel_exporter.export_to_excel(all_data, max_dates, max_prices)
                print(f"\nProcessed {base_name} -> {output_file}")

        elif choice == '2':
            if menu_handler.show_settings_menu() == 'edit':
                new_keywords = menu_handler.edit_keywords()
                if new_keywords:
                    menu_handler.settings["title_keywords"] = new_keywords
                    menu_handler.settings_manager.save_settings(menu_handler.settings)

        elif choice == '3':
            print("\nGoodbye!")
            exit()

if __name__ == "__main__":
    main()