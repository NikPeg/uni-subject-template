#!/usr/bin/env python3
"""
Скрипт для конвертации DOCX файлов в TXT с анализом изображений через OpenRouter API.
Сохраняет правильную последовательность текста и изображений.
"""

import os
import re
import sys
import base64
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Tuple, Dict
import requests
from dotenv import load_dotenv
try:
    from docx import Document
    from docx.document import Document as DocumentType
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("Предупреждение: python-docx не установлен. Будет использован базовый парсинг XML.")

# Загружаем переменные окружения
load_dotenv()

OPENROUTER_API_KEY = os.getenv("LLM_TOKEN", "").strip('"')
VISION_MODEL = os.getenv("VISION_MODEL", "google/gemini-2.0-flash-001")
OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"


def extract_images_from_docx(docx_path: str) -> dict:
    """Извлекает все изображения из DOCX файла."""
    images = {}
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        for file_info in zip_ref.namelist():
            if file_info.startswith('word/media/'):
                images[file_info] = zip_ref.read(file_info)
    return images


def get_image_base64(image_data: bytes) -> str:
    """Конвертирует изображение в base64 строку."""
    return base64.b64encode(image_data).decode('utf-8')


def analyze_image_with_api(image_data: bytes, image_name: str) -> str:
    """Отправляет изображение в OpenRouter API для анализа."""
    if not OPENROUTER_API_KEY:
        return f"[Изображение: {image_name} - API токен не найден]"
    
    try:
        # Конвертируем в base64
        image_base64 = get_image_base64(image_data)
        
        # Определяем MIME тип по расширению
        ext = Path(image_name).suffix.lower()
        mime_types = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.webp': 'image/webp'
        }
        mime_type = mime_types.get(ext, 'image/png')
        
        # Формируем запрос к OpenRouter API
        headers = {
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com",  # Опционально, для отслеживания
        }
        
        payload = {
            "model": VISION_MODEL,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Опиши подробно это изображение. Если на нем есть текст, перепиши его полностью. Если это диаграмма, график или схема, опиши их структуру и содержание."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{image_base64}"
                            }
                        }
                    ]
                }
            ],
            "max_tokens": 4000
        }
        
        response = requests.post(OPENROUTER_API_URL, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        
        result = response.json()
        if 'choices' in result and len(result['choices']) > 0:
            return result['choices'][0]['message']['content']
        else:
            return f"[Изображение: {image_name} - Не удалось получить описание]"
            
    except Exception as e:
        return f"[Изображение: {image_name} - Ошибка анализа: {str(e)}]"


def parse_docx_structure(docx_path: str) -> Tuple[List[Tuple[str, str]], Dict[str, bytes]]:
    """
    Парсит DOCX файл и возвращает список элементов (текст или изображение) в правильном порядке.
    Возвращает (список кортежей: ('text', content) или ('image', image_path_in_zip), словарь изображений)
    """
    elements = []
    images = extract_images_from_docx(docx_path)
    
    if HAS_DOCX:
        # Используем python-docx для более надежного парсинга
        try:
            doc = Document(docx_path)
            
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                # Читаем relationships для связи ID изображений с путями
                rels_xml = zip_ref.read('word/_rels/document.xml.rels')
                rels_root = ET.fromstring(rels_xml)
                rels_ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                
                rel_to_image = {}
                for rel in rels_root.findall('.//r:Relationship', rels_ns):
                    target = rel.get('Target', '')
                    if target.startswith('media/'):
                        rel_to_image[rel.get('Id')] = f'word/{target}'
                
                # Парсим XML напрямую для поиска изображений в правильном порядке
                document_xml = zip_ref.read('word/document.xml')
                root = ET.fromstring(document_xml)
                ns = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
                }
                
                # Обрабатываем все параграфы в порядке их появления
                for paragraph in root.findall('.//w:p', ns):
                    # Извлекаем весь текст параграфа
                    text_parts = []
                    for t in paragraph.findall('.//w:t', ns):
                        if t.text:
                            text_parts.append(t.text)
                    
                    # Проверяем наличие рисунков в параграфе
                    drawings = paragraph.findall('.//w:drawing', ns)
                    image_found = False
                    
                    if drawings:
                        for drawing in drawings:
                            # Ищем blip (изображение)
                            blip = drawing.find('.//a:blip', ns)
                            if blip is not None:
                                r_embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                if r_embed and r_embed in rel_to_image:
                                    image_path = rel_to_image[r_embed]
                                    if image_path in images:
                                        # Добавляем текст перед изображением, если есть
                                        if text_parts:
                                            text_content = ''.join(text_parts).strip()
                                            if text_content:
                                                elements.append(('text', text_content))
                                        elements.append(('image', image_path))
                                        image_found = True
                                        text_parts = []  # Очищаем, так как уже добавили
                                        break
                    
                    # Добавляем текст, если изображение не найдено или текст остался
                    if not image_found and text_parts:
                        text_content = ''.join(text_parts).strip()
                        if text_content:
                            elements.append(('text', text_content))
            
            return elements, images
        except Exception as e:
            print(f"  Предупреждение: ошибка при использовании python-docx: {e}")
            print("  Переключение на базовый XML парсинг...")
    
    # Базовый XML парсинг (fallback)
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        document_xml = zip_ref.read('word/document.xml')
        root = ET.fromstring(document_xml)
        
        ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
        
        # Читаем relationships
        try:
            rels_xml = zip_ref.read('word/_rels/document.xml.rels')
            rels_root = ET.fromstring(rels_xml)
            rels_ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            
            rel_to_image = {}
            for rel in rels_root.findall('.//r:Relationship', rels_ns):
                target = rel.get('Target', '')
                if target.startswith('media/'):
                    rel_to_image[rel.get('Id')] = f'word/{target}'
        except:
            rel_to_image = {}
        
        # Обрабатываем параграфы
        for paragraph in root.findall('.//w:p', ns):
            text_parts = []
            for t in paragraph.findall('.//w:t', ns):
                if t.text:
                    text_parts.append(t.text)
            
            drawings = paragraph.findall('.//w:drawing', ns)
            image_found = False
            
            if drawings:
                for drawing in drawings:
                    blip = drawing.find('.//a:blip', ns)
                    if blip is not None:
                        r_embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if r_embed and r_embed in rel_to_image:
                            image_path = rel_to_image[r_embed]
                            if image_path in images:
                                if text_parts:
                                    text_content = ''.join(text_parts).strip()
                                    if text_content:
                                        elements.append(('text', text_content))
                                elements.append(('image', image_path))
                                image_found = True
                                text_parts = []
                                break
            
            if not image_found and text_parts:
                text_content = ''.join(text_parts).strip()
                if text_content:
                    elements.append(('text', text_content))
    
    return elements, images


def convert_docx_to_txt(docx_path: str, output_path: str = None) -> str:
    """Конвертирует DOCX файл в TXT с анализом изображений."""
    if output_path is None:
        output_path = str(Path(docx_path).with_suffix('.txt'))
    
    print(f"Обработка: {docx_path}")
    
    # Парсим структуру документа
    elements, images = parse_docx_structure(docx_path)
    
    if not elements:
        print(f"  Предупреждение: не удалось извлечь содержимое из {docx_path}")
        return output_path
    
    # Формируем итоговый текст
    result_lines = []
    image_count = 0
    
    for element_type, content in elements:
        if element_type == 'text':
            result_lines.append(content)
            result_lines.append('')  # Пустая строка после текста
        elif element_type == 'image':
            image_count += 1
            image_name = Path(content).name
            print(f"  Анализ изображения {image_count}: {image_name}")
            
            if content in images:
                image_description = analyze_image_with_api(images[content], image_name)
                result_lines.append(f"\n[ИЗОБРАЖЕНИЕ {image_count}: {image_name}]")
                result_lines.append(image_description)
                result_lines.append('')  # Пустая строка после описания
            else:
                result_lines.append(f"\n[ИЗОБРАЖЕНИЕ {image_count}: {image_name} - не найдено]")
                result_lines.append('')
    
    # Сохраняем результат
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(result_lines))
    
    print(f"  Готово: {output_path} ({image_count} изображений обработано)")
    return output_path


def main():
    """Основная функция."""
    # Обрабатываем аргументы командной строки
    if len(sys.argv) > 1:
        target_path = Path(sys.argv[1])
    else:
        # Если аргументов нет, используем текущую директорию
        target_path = Path('.')
    
    # Определяем, что передано: файл или папка
    if not target_path.exists():
        print(f"ОШИБКА: Путь '{target_path}' не существует!")
        return
    
    # Собираем список файлов для обработки
    docx_files = []
    
    if target_path.is_file():
        # Если передан файл, обрабатываем только его
        if target_path.suffix.lower() == '.docx':
            docx_files = [target_path]
        else:
            print(f"ОШИБКА: '{target_path}' не является DOCX файлом!")
            return
    elif target_path.is_dir():
        # Если передана папка, находим все DOCX файлы в ней
        docx_files = list(target_path.glob('*.docx'))
    else:
        print(f"ОШИБКА: '{target_path}' не является файлом или папкой!")
        return
    
    if not docx_files:
        if target_path.is_file():
            print(f"Файл '{target_path}' не найден или не является DOCX файлом.")
        else:
            print(f"DOCX файлы не найдены в '{target_path}'.")
        return
    
    print(f"Найдено {len(docx_files)} DOCX файлов для обработки.\n")
    
    # Проверяем наличие токена
    if not OPENROUTER_API_KEY:
        print("ОШИБКА: LLM_TOKEN не найден в .env файле!")
        return
    
    print(f"Используется модель: {VISION_MODEL}\n")
    
    # Обрабатываем каждый файл
    for docx_file in docx_files:
        # Пропускаем файлы, для которых уже существует .txt файл
        txt_file = docx_file.with_suffix('.txt')
        if txt_file.exists():
            print(f"Пропуск {docx_file.name} - файл {txt_file.name} уже существует.\n")
            continue
        
        try:
            convert_docx_to_txt(str(docx_file))
            print()
        except Exception as e:
            print(f"ОШИБКА при обработке {docx_file}: {str(e)}\n")
    
    print("Обработка завершена!")


if __name__ == '__main__':
    main()

