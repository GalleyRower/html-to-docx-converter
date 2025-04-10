import os
import sys
import re
from bs4 import BeautifulSoup, NavigableString
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_UNDERLINE

def sanitize_text(text):
    """Удаляет или заменяет проблемные символы."""
    text = text.replace('\x0b', '')  # Удаляем вертикальную табуляцию
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)  # Удаляем другие control characters
    return text

def html_to_docx(text_file_path, docx_file_path):
    """
    Преобразует HTML-код из текстового файла в DOCX-файл, сохраняя форматирование.
    """
    try:
        with open(text_file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        soup = BeautifulSoup(html_content, 'html.parser')
        document = Document()

        # --- Определение стилей ---
        styles = document.styles
        # Заголовок по центру
        title_style = styles.add_style('TitleCenter', WD_STYLE_TYPE.PARAGRAPH)
        title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_style.font.size = Pt(20) # Увеличен размер для наглядности
        title_style.font.bold = True

        # Стиль для ссылок
        link_style = styles.add_style('Link', WD_STYLE_TYPE.CHARACTER)
        link_style.font.underline = True
        link_style.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)  # Синий цвет для ссылок

        def process_element(element, paragraph=None):
            """Рекурсивно обрабатывает HTML-элементы и добавляет их в DOCX."""
            if isinstance(element, NavigableString):
                text = sanitize_text(str(element)) # Очищаем текст
                if paragraph is not None and text: # Добавляем только непустой текст
                    paragraph.add_run(text)
                return

            for child in element.contents:
                if child.name == 'h1':
                    p = document.add_paragraph(style='TitleCenter') # Используем style
                    process_element(child, p)
                elif child.name == 'h2':
                    p = document.add_paragraph(style='Heading 2')
                    process_element(child, p)
                elif child.name == 'h3':
                    p = document.add_paragraph(style='Heading 3')
                    process_element(child, p)
                elif child.name == 'p':
                    p = document.add_paragraph()
                    process_element(child, p)
                elif child.name == 'b' or child.name == 'strong':
                    if paragraph is not None:
                        run = paragraph.add_run()
                        run.text = sanitize_text(child.text)
                        run.bold = True
                elif child.name == 'i' or child.name == 'em':
                    if paragraph is not None:
                        run = paragraph.add_run()
                        run.text = sanitize_text(child.text)
                        run.italic = True
                elif child.name == 'u': # Подчеркнутый текст
                    if paragraph is not None:
                        run = paragraph.add_run()
                        run.text = sanitize_text(child.text)
                        run.underline = True
                elif child.name == 'br':
                    if paragraph is not None:
                        run = paragraph.add_run()
                        run.add_break(WD_BREAK.LINE)
                elif child.name == 'a':
                    if paragraph is not None:
                        run = paragraph.add_run(sanitize_text(child.text))
                        run.style = 'Link' #Применяем style ссылок
                        # В DOCX нет прямой поддержки ссылок, можно добавить URL в конце
                        # run.text += f" ({child['href']})" # Убрал отображение ссылки рядом с текстом, чтобы не засорять документ
                elif child.name == 'ul':
                    for li in child.find_all('li'):
                        p = document.add_paragraph(f"• {sanitize_text(li.text)}")  # Используем маркер списка
                elif child.name == 'ol':
                    i = 1
                    for li in child.find_all('li'):
                        p = document.add_paragraph(f"{i}. {sanitize_text(li.text)}")  # Нумерованный список
                        i += 1
                elif child.name == 'table':
                    #  Обработка таблиц (улучшенная)
                    table = document.add_table(rows=0, cols=0)
                    for row in child.find_all('tr'):
                        cells = row.find_all('td')
                        if not cells:
                            cells = row.find_all('th')  # Заголовки таблицы
                        docx_row = table.add_row()
                        num_cells = len(cells)
                        max_cells = len(docx_row.cells)  # Максимальное кол-во ячеек в строке DOCX
                        for i in range(min(num_cells, max_cells)): # Ограничиваем кол-во ячеек
                            try:
                                cell_text = sanitize_text(cells[i].text)
                                docx_row.cells[i].text = cell_text
                            except IndexError:
                                print(f"Индекс ячейки {i} вне диапазона в строке таблицы") # Добавляем отладочное сообщение
                                pass # Пропускаем проблемную ячейку
                            except Exception as e:
                                print(f"Ошибка при обработке ячейки таблицы: {e}")
                                pass # Пропускаем проблемную ячейку
                else:
                    # Просто добавляем текст для неизвестных тегов
                    if paragraph is not None:
                        process_element(child, paragraph)

        process_element(soup.body if soup.body else soup)  # Начинаем обработку с <body>

        document.save(docx_file_path)
        return True

    except Exception as e:
        print(f"Ошибка при преобразовании {text_file_path}: {e}")
        return False

def process_folder(input_folder, output_folder):
    """
    Обрабатывает все текстовые файлы в указанной входной папке и сохраняет
    DOCX-файлы в выходной папке.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    total_files = 0
    converted_files = 0
    failed_files = []

    for filename in os.listdir(input_folder):
        if filename.endswith(".txt"):  # Проверяем расширение .txt
            total_files += 1
            text_file_path = os.path.join(input_folder, filename)
            docx_file_path = os.path.join(output_folder, os.path.splitext(filename)[0] + ".docx")  # Заменяем расширение

            try: # Общая обработка исключений для файла
                if html_to_docx(text_file_path, docx_file_path):
                    converted_files += 1
                    print(f"Успешно преобразован: {text_file_path} -> {docx_file_path}")
                else:
                    print(f"Ошибка при преобразовании {text_file_path}")
                    failed_files.append(filename)
            except Exception as e:
                print(f"Непредвиденная ошибка при обработке {text_file_path}: {e}")
                failed_files.append(filename)

    return total_files, converted_files, failed_files  # Возвращаем failed_files

def main():
    """
    Основная функция, которая получает пути входной и выходной папок
    из аргументов командной строки.
    """
    if len(sys.argv) != 3:
        print("Использование: python script.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]

    if not os.path.exists(input_folder):
        print(f"Ошибка: Входная папка '{input_folder}' не найдена.")
        sys.exit(1)

    total_files, converted_files, failed_files = process_folder(input_folder, output_folder)

    print("\n-----------------------------------")
    print("Обработка завершена.")
    print(f"Всего файлов найдено: {total_files}")
    print(f"Успешно преобразовано: {converted_files}")
    print(f"Не удалось преобразовать: {total_files - converted_files}")
    if failed_files:
        print("\nНе удалось преобразовать следующие файлы:")
        for filename in failed_files:
            print(f"- {filename}")
    print("-----------------------------------")

if __name__ == "__main__":
    main()
