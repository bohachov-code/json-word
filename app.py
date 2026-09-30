import json
import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def convert_json_to_docx(json_data):
    doc = Document()

    # Заголовок документа
    heading = doc.add_heading('ІНДИВІДУАЛЬНИЙ ПЛАН РОБОТИ', level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Основні дані викладача
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.25

    p.add_run("Факультет: ").bold = True
    p.add_run(f"{json_data.get('fac', '—')}\n")[cite: 1]

    p.add_run("Кафедра: ").bold = True
    p.add_run(f"{json_data.get('kaf', '—')}\n")[cite: 1]

    p.add_run("ПІБ: ").bold = True
    p.add_run(f"{json_data.get('fio', '—')}\n")[cite: 1]

    p.add_run("Посада: ").bold = True
    p.add_run(f"{json_data.get('pos', '—')}\n")[cite: 1]

    deg = json_data.get('deg')[cite: 1]
    if deg:
        p.add_run("Науковий ступінь: ").bold = True
        p.add_run(f"{deg}\n")[cite: 1]

    p.add_run("Навчальний рік: ").bold = True
    p.add_run(f"{json_data.get('yr', '—')}")[cite: 1]

    doc.add_paragraph()  # Відступ

    # Розділ "План робіт"
    h2 = doc.add_heading('План робіт', level=2)

    plan = json_data.get('plan', [])[cite: 1]

    if plan:
        # Створюємо таблицю з 5 колонками
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'

        # Заголовки стовпців
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Код (ID)'
        hdr_cells[1].text = 'Обсяг (годин / шт)'
        hdr_cells[2].text = 'Термін виконання'
        hdr_cells[3].text = 'Результат / Назва'
        hdr_cells[4].text = 'Примітка'

        # Заповнюємо рядки
        for item in plan:
            row_cells = table.add_row().cells
            row_cells[0].text = str(item.get('id', ''))[cite: 1]
            row_cells[1].text = str(item.get('qty', 0))[cite: 1]
            row_cells[2].text = str(item.get('sem', ''))[cite: 1]
            row_cells[3].text = str(item.get('res', ''))[cite: 1]
            row_cells[4].text = str(item.get('cmt', ''))[cite: 1]

    # Збереження документа в пам'ять (байтовий потік)
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


# --- Інтерфейс Streamlit ---
st.set_page_config(page_title="Конвертер Індивідуальних Планів", page_icon="📄")

st.title("📄 Генератор Індивідуального Плану (Word)")
st.write("Завантажте ваші `.json` файли, щоб отримати готовий `.docx` документ.")

uploaded_file = st.file_uploader("Оберіть файл JSON", type=["json"])

if uploaded_file is not None:
    try:
        json_data = json.load(uploaded_file)
        docx_file = convert_json_to_docx(json_data)

        # Замінюємо розширення .json на .docx для вихідного файлу
        output_filename = uploaded_file.name.rsplit('.', 1)[0] + '.docx'

        st.success("Документ успішно сформовано!")
        st.download_button(
            label="💾 Завантажити Word (.docx)",
            data=docx_file,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Помилка зчитування або обробки JSON: {e}")