import PyPDF2
import pandas as pd
import docx
import os
from pathlib import Path
import re
from odf import text, teletype
from odf.opendocument import load
import json
from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET
from striprtf.striprtf import rtf_to_text
import zipfile
import rarfile

class DocumentParser:
    def __init__(self, file_paths):
        self.file_paths = [Path(fp) for fp in file_paths]

    def search_in_pdf(self, file_path, search_terms):
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                num_pages = len(pdf_reader.pages)
                results = {}
                for page_num in range(num_pages):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text().lower()
                    for term in search_terms:
                        term = term.lower()
                        matches = [(m.start(), m.end()) for m in re.finditer(r'\b' + re.escape(term) + r'\b', text)]
                        if matches:
                            if term not in results:
                                results[term] = []
                            results[term].append({'page': page_num + 1, 'count': len(matches), 'positions': matches[:5]})
                return {'type': 'pdf', 'pages': num_pages, 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге PDF: {str(e)}"}

    def search_in_excel(self, file_path, search_terms):
        try:
            excel_file = pd.ExcelFile(file_path)
            results = {}
            total_rows = 0
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                total_rows += len(df)
                for term in search_terms:
                    term = term.strip()
                    try:
                        term_float = float(term.replace(',', '.'))
                    except ValueError:
                        term_float = None
                    term_lower = term.lower()
                    matches = {}
                    for col in df.columns:
                        for idx, value in df[col].items():
                            if pd.notna(value):
                                if isinstance(value, (int, float)):
                                    if term_float is not None and float(value) == term_float:
                                        if term_lower not in matches:
                                            matches[term_lower] = []
                                        matches[term_lower].append({'sheet': sheet_name, 'row': idx + 2, 'column': col, 'cell_value': value})
                                else:
                                    value_str = str(value).lower().strip()
                                    if value_str == term_lower:
                                        if term_lower not in matches:
                                            matches[term_lower] = []
                                        matches[term_lower].append({'sheet': sheet_name, 'row': idx + 2, 'column': col, 'cell_value': value})
                    if matches:
                        if term_lower not in results:
                            results[term_lower] = []
                        results[term_lower].extend(matches[term_lower])
            return {'type': 'excel', 'sheets': excel_file.sheet_names, 'total_rows': total_rows, 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге Excel: {str(e)}"}

    def search_in_docx(self, file_path, search_terms):
        try:
            doc = docx.Document(file_path)
            text = "\n".join([para.text for para in doc.paragraphs]).lower()
            results = {}
            for term in search_terms:
                term = term.lower()
                matches = [(m.start(), m.end()) for m in re.finditer(r'\b' + re.escape(term) + r'\b', text)]
                if matches:
                    results[term] = [{'count': len(matches), 'positions': matches[:5]}]
            return {'type': 'docx', 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге DOCX: {str(e)}"}

    def search_in_txt(self, file_path, search_terms):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read().lower()
            results = {}
            for term in search_terms:
                term = term.lower()
                matches = [(m.start(), m.end()) for m in re.finditer(r'\b' + re.escape(term) + r'\b', text)]
                if matches:
                    results[term] = [{'count': len(matches), 'positions': matches[:5]}]
            return {'type': 'txt', 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге TXT: {str(e)}"}

    def search_in_csv(self, file_path, search_terms):
        try:
            df = pd.read_csv(file_path)
            results = {}
            total_rows = len(df)
            for term in search_terms:
                term = term.strip()
                try:
                    term_float = float(term.replace(',', '.'))
                except ValueError:
                    term_float = None
                term_lower = term.lower()
                matches = {}
                for col in df.columns:
                    for idx, value in df[col].items():
                        if pd.notna(value):
                            if isinstance(value, (int, float)):
                                if term_float is not None and float(value) == term_float:
                                    if term_lower not in matches:
                                        matches[term_lower] = []
                                    matches[term_lower].append({'row': idx + 2, 'column': col, 'cell_value': value})
                            else:
                                value_str = str(value).lower().strip()
                                if value_str == term_lower:
                                    if term_lower not in matches:
                                        matches[term_lower] = []
                                    matches[term_lower].append({'row': idx + 2, 'column': col, 'cell_value': value})
                if matches:
                    if term_lower not in results:
                        results[term_lower] = []
                    results[term_lower].extend(matches[term_lower])
            return {'type': 'csv', 'total_rows': total_rows, 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге CSV: {str(e)}"}

    def search_in_odt(self, file_path, search_terms):
        try:
            doc = load(file_path)
            text_content = "\n".join([teletype.extractText(p) for p in doc.getElementsByType(text.P)]).lower()
            results = {}
            for term in search_terms:
                term = term.lower()
                matches = [(m.start(), m.end()) for m in re.finditer(r'\b' + re.escape(term) + r'\b', text_content)]
                if matches:
                    results[term] = [{'count': len(matches), 'positions': matches[:5]}]
            return {'type': 'odt', 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге ODT: {str(e)}"}

    def search_in_json(self, file_path, search_terms):
        # ... (без изменений)
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
            results = {}
            def search_recursive(obj, term, path=""):
                matches = []
                if isinstance(obj, dict):
                    for key, value in obj.items():
                        new_path = f"{path}.{key}" if path else key
                        matches.extend(search_recursive(value, term, new_path))
                elif isinstance(obj, list):
                    for i, item in enumerate(obj):
                        new_path = f"{path}[{i}]"
                        matches.extend(search_recursive(item, term, new_path))
                elif isinstance(obj, (str, int, float)):
                    value_str = str(obj).lower().strip()
                    if value_str == term:
                        matches.append({'path': path, 'value': obj})
                return matches
            for term in search_terms:
                term = term.lower().strip()
                try:
                    term_float = float(term.replace(',', '.'))
                except ValueError:
                    term_float = None
                matches = search_recursive(data, term if term_float is None else str(term_float))
                if matches:
                    results[term] = matches
            return {'type': 'json', 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге JSON: {str(e)}"}

    def search_in_html(self, file_path, search_terms):
        # ... (без изменений)
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                soup = BeautifulSoup(file, 'html.parser')
                text = soup.get_text().lower()
            results = {}
            for term in search_terms:
                term = term.lower()
                matches = [(m.start(), m.end()) for m in re.finditer(r'\b' + re.escape(term) + r'\b', text)]
                if matches:
                    results[term] = [{'count': len(matches), 'positions': matches[:5]}]
            return {'type': 'html', 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге HTML: {str(e)}"}

    def search_in_xml(self, file_path, search_terms):
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            results = {}
            def search_recursive(elem, term, path=""):
                matches = []
                if elem.text:
                    text = elem.text.lower().strip()
                    if text == term:
                        matches.append({'path': f"{path}.text", 'value': elem.text})
                for key, value in elem.attrib.items():
                    if value.lower().strip() == term:
                        matches.append({'path': f"{path}.@{key}", 'value': value})
                for child in elem:
                    new_path = f"{path}/{child.tag}" if path else child.tag
                    matches.extend(search_recursive(child, term, new_path))
                return matches
            for term in search_terms:
                term = term.lower().strip()
                matches = search_recursive(root, term)
                if matches:
                    results[term] = matches
            return {'type': 'xml', 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге XML: {str(e)}"}

    def search_in_rtf(self, file_path, search_terms):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                rtf_content = file.read()
            text = rtf_to_text(rtf_content).lower()
            results = {}
            for term in search_terms:
                term = term.lower()
                matches = [(m.start(), m.end()) for m in re.finditer(r'\b' + re.escape(term) + r'\b', text)]
                if matches:
                    results[term] = [{'count': len(matches), 'positions': matches[:5]}]
            return {'type': 'rtf', 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге RTF: {str(e)}"}

    def search_in_md(self, file_path, search_terms):
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read().lower()
            results = {}
            for term in search_terms:
                term = term.lower()
                matches = [(m.start(), m.end()) for m in re.finditer(r'\b' + re.escape(term) + r'\b', text)]
                if matches:
                    results[term] = [{'count': len(matches), 'positions': matches[:5]}]
            return {'type': 'md', 'search_results': results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге Markdown: {str(e)}"}

    def search_in_zip(self, file_path, search_terms):
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                temp_dir = Path('temp_extract_zip')
                temp_dir.mkdir(exist_ok=True)
                zip_ref.extractall(temp_dir)
                nested_results = {}
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        nested_file_path = Path(root) / file
                        nested_result = self.parse_and_search_single(nested_file_path, search_terms)
                        if nested_result and 'error' not in nested_result:
                            nested_results[nested_file_path.relative_to(temp_dir).as_posix()] = nested_result
                for root, dirs, files in os.walk(temp_dir, topdown=False):
                    for file in files:
                        os.remove(os.path.join(root, file))
                    for dir in dirs:
                        os.rmdir(os.path.join(root, dir))
                os.rmdir(temp_dir)
                return {'type': 'zip', 'search_results': nested_results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге ZIP: {str(e)}"}

    def search_in_rar(self, file_path, search_terms):
        try:
            with rarfile.RarFile(file_path, 'r') as rar_ref:
                temp_dir = Path('temp_extract_rar')
                temp_dir.mkdir(exist_ok=True)
                rar_ref.extractall(temp_dir)
                nested_results = {}
                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        nested_file_path = Path(root) / file
                        nested_result = self.parse_and_search_single(nested_file_path, search_terms)
                        if nested_result and 'error' not in nested_result:
                            nested_results[nested_file_path.relative_to(temp_dir).as_posix()] = nested_result
                for root, dirs, files in os.walk(temp_dir, topdown=False):
                    for file in files:
                        os.remove(os.path.join(root, file))
                    for dir in dirs:
                        os.rmdir(os.path.join(root, dir))
                os.rmdir(temp_dir)
                return {'type': 'rar', 'search_results': nested_results}
        except Exception as e:
            return {'error': f"Ошибка при парсинге RAR: {str(e)}"}

    def parse_and_search_single(self, file_path, search_terms):
        ext = file_path.suffix.lower()
        if ext == '.pdf':
            return self.search_in_pdf(file_path, search_terms)
        elif ext in ['.xls', '.xlsx']:
            return self.search_in_excel(file_path, search_terms)
        elif ext == '.docx':
            return self.search_in_docx(file_path, search_terms)
        elif ext == '.txt':
            return self.search_in_txt(file_path, search_terms)
        elif ext == '.csv':
            return self.search_in_csv(file_path, search_terms)
        elif ext == '.odt':
            return self.search_in_odt(file_path, search_terms)
        elif ext == '.json':
            return self.search_in_json(file_path, search_terms)
        elif ext == '.html':
            return self.search_in_html(file_path, search_terms)
        elif ext == '.xml':
            return self.search_in_xml(file_path, search_terms)
        elif ext == '.rtf':
            return self.search_in_rtf(file_path, search_terms)
        elif ext == '.md':
            return self.search_in_md(file_path, search_terms)
        elif ext == '.zip':
            return self.search_in_zip(file_path, search_terms)
        elif ext == '.rar':
            return self.search_in_rar(file_path, search_terms)
        else:
            return {'error': 'Неподдерживаемый формат файла'}

    def parse_and_search(self, search_terms):
        all_results = {}
        for file_path in self.file_paths:
            if not file_path.exists():
                all_results[file_path.name] = {'error': 'Файл не найден'}
                continue
            all_results[file_path.name] = self.parse_and_search_single(file_path, search_terms)
        return all_results