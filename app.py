import streamlit as st
import pandas as pd
import io
import os
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont 

EXCEL_FILE = 'Бочки 3-я фура.xlsx'
WORD_OUTPUT = 'Labels_150x100.docx'
PDF_OUTPUT = 'Labels_150x100.pdf'

LABEL_W_MM = 150
LABEL_H_MM = 100
LABEL_W, LABEL_H = LABEL_W_MM * mm, LABEL_H_MM * mm

def setup_pdf_fonts():
    """Registers fonts. If Arial.ttf isn't in your folder, it uses Helvetica."""
    try:
        pdfmetrics.registerFont(TTFont('ArialCustom', 'Arial.ttf'))
        pdfmetrics.registerFont(TTFont('ArialBoldCustom', 'Arial-Bold.ttf'))
        return 'ArialCustom', 'ArialBoldCustom'
    except:
        return 'Helvetica', 'Helvetica-Bold'



def create_word_file(data):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Mm(LABEL_W_MM), Mm(LABEL_H_MM)
    section.top_margin, section.bottom_margin = Mm(10), Mm(10)
    section.left_margin, section.right_margin = Mm(12), Mm(12)

    for i, row in data.iterrows():
        if i > 0: doc.add_page_break()
        table = doc.add_table(rows=1, cols=3)
        table.width = Mm(LABEL_W_MM - 24)
        
        mid = table.cell(0, 1).paragraphs[0]
        mid.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = mid.add_run("ПИЦЦА СОУС")
        r.bold, r.font.size = True, Pt(8)

        right = table.cell(0, 2).paragraphs[0]
        right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = right.add_run("(ПРОДУКТ ПАСТЕРИЗОВАН)")
        r.bold, r.font.size = True, Pt(6)

        def tight_p(text):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.3
            run = p.add_run(text)
            run.bold, run.font.size = True, Pt(11)

        p_date = pd.to_datetime(row['Дата Производства']).strftime('%d.%m.%Y')
        e_date = pd.to_datetime(row['Годен до']).strftime('%d.%m.%Y')
        ph = f"{float(row['PH']):.2f}"

        tight_p(f"Номер партии: {row['Номер Партии']}  /  BRIX: {row['BRIX']}%  /  PH: {ph}")
        tight_p(f"Вес Нетто: {row['Нетто соуса']}  /  Вес Брутто: {row['Брутто бочек']}")
        tight_p(f"Изготовлено: {p_date}  /  Годен до: {e_date}")

        legal = doc.add_paragraph()
        legal.paragraph_format.line_spacing = 1.3
        run1 = legal.add_run(
            "--------------------------------------------------------------------------------\n"
            "Состав: Томаты, соль пищевая, регулятор кислотности (лимонная кислота).\n"
            "Пищевая ценность на 100 г. продукта: углеводы - 9,0г, жиры - 0,0г\n"
            "энергетическая ценность - 36,0 ккал/153,0 КДж.\n"
            "Хранить в сухом, прохладном месте при температуре от +0°С до +25°C и\n"
            "относительной влажности воздуха не более 75%. После вскрытия хранить при\n"
            "температуре от +0°С до +5°С не более 24 часа.\n"
            "Срок годности: 12 месяцев с даты изготовления.\n"
            'Изготовитель: ООО "Эребуни Трейд Груп".\n'
            "Юридический адрес: РА, г. Ереван, ул. Аветисян, дом 23.\n"
            "Адрес производства: Республика Армения, Араратский рег., г. Арташат,\n"
            "ул. Самвела Акопяна, стр. 173.\n"
            "Тел: (+374) 77-733-388 Email: erebunit@gmail.com\n"
            "ТУ AM 51192101. 9183-2023\n"
        )
        run1.font.size = Pt(9)

        run2 = legal.add_run("Произведено в Республике Армения.")
        run2.bold = True
        run2.font.size = Pt(9)

    target = io.BytesIO()
    doc.save(target)
    return target.getvalue()


def create_pdf_file(data):
    f_reg, f_bold = setup_pdf_fonts()
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(LABEL_W, LABEL_H))
    
    MARGIN_X, MARGIN_TOP = 5 * mm, 7 * mm
    pallet_counter = 1

    for index, row in data.iterrows():
        c.setFont(f_bold, 12)
        c.drawCentredString(LABEL_W / 2, LABEL_H - MARGIN_TOP, "ПИЦЦА СОУС")
        
        c.setFont(f_bold, 10)
        c.drawRightString(LABEL_W - MARGIN_X, LABEL_H - MARGIN_TOP, "(ПРОДУКТ ПАСТЕРИЗОВАН)")

        p_date = pd.to_datetime(row['Дата Производства']).strftime('%d.%m.%Y')
        e_date = pd.to_datetime(row['Годен до']).strftime('%d.%m.%Y')
        ph = f"{float(row['PH']):.2f}"

        y = LABEL_H - MARGIN_TOP - 5 * mm
        c.setFont(f_bold, 10)

        c.drawString(MARGIN_X, y,
            f"Номер партии: {row['Номер Партии']}  /  BRIX: {row['BRIX']}%  /  PH: {ph}")
        y -= 4 * mm
        c.drawString(MARGIN_X, y,
            f"Вес Нетто: {row['Нетто соуса']}  /  Вес Брутто: {row['Брутто бочек']}")
        y -= 4 * mm
        c.drawString(MARGIN_X, y,
            f"Изготовлено: {p_date}  /  Годен до: {e_date}")

        c.line(MARGIN_X, y - 3 * mm, LABEL_W - MARGIN_X, y - 3 * mm)

        c.setFont(f_reg, 10)
        y -= 10 * mm

        text = c.beginText(MARGIN_X, y)
        text.setLeading(13)  

        for line in [
            "Состав: Томаты, соль пищевая, регулятор кислотности (лимонная кислота).",
            "Пищевая ценность на 100 г. продукта: углеводы - 9,0г, жиры - 0,0г",
            "энергетическая ценность - 36,0 ккал/153,0 КДж.",
            "Хранить в сухом, прохладном месте при температуре от +0°С до +25°C и",
            "относительной влажности воздуха не более 75%. После вскрытия хранить при",
            "температуре от +0°С до +5°С не более 24 часа.",
            "Срок годности: 12 месяцев с даты изготовления.",
            'Изготовитель: ООО "Эребуни Трейд Груп".',
            "Юридический адрес: РА, г. Ереван, ул. Аветисян, дом 23.",
            "Адрес производства: Республика Армения, Араратский рег., г. Арташат,",
            "ул. Самвела Акопяна, стр. 173.",
            "Тел: (+374) 77-733-388 Email: erebunit@gmail.com",
            "",
            "ТУ AM 51192101. 9183-2023",
        ]:
            text.textLine(line)

        c.drawText(text)

        c.setFont(f_bold, 10)
        c.drawString(MARGIN_X, text.getY() - 3, "Произведено в Республике Армения.")

        LOGO_SIZE = 10 * mm
        logo_gap = 1 * mm
        LOGO_Y = 3 * mm

        if os.path.exists("logo_right.png"):
            c.drawImage("logo_right.png",
                        LABEL_W - LOGO_SIZE - MARGIN_X,
                        LOGO_Y,
                        width=LOGO_SIZE, height=LOGO_SIZE,
                        preserveAspectRatio=True, mask='auto')

        if os.path.exists("logo_left.png"):
            c.drawImage("logo_left.png",
                        LABEL_W - 2*LOGO_SIZE - MARGIN_X - logo_gap,
                        LOGO_Y,
                        width=LOGO_SIZE, height=LOGO_SIZE,
                        preserveAspectRatio=True, mask='auto')

        c.showPage()

        if (index + 1) % 4 == 0:
            c.setFont(f_bold, 28)
            c.drawCentredString(LABEL_W / 2, 75 * mm, f"ПАЛЕТА \u2116 {pallet_counter:02d}")
            c.setFont(f_bold, 24)
            val_net = row['Нетто соуса на паллете']
            val_gross = row['Брутто паллета']
            
            net_str = f"{float(val_net):.1f}".replace('.', ',') if pd.notna(val_net) else "0,0"
            gross_str = f"{float(val_gross):.1f}".replace('.', ',') if pd.notna(val_gross) else "0,0"
            
            LEFT_ALIGN_X = 25 * mm 
            c.drawString(LEFT_ALIGN_X, 45 * mm, f"Вес Нетто  -  {net_str}")
            c.drawString(LEFT_ALIGN_X, 20 * mm, f"Вес Брутто  -  {gross_str}")
            pallet_counter += 1
            c.showPage()

    c.save()
    return buffer.getvalue()

<<<<<<< HEAD
st.set_page_config(
    page_title="Erebuni Label Gen", 
    page_icon="📦", 
    layout="centered" # Changed to centered for a focused look
)
=======
st.set_page_config(page_title="Erebuni Label Gen", page_icon="📦")
st.title("Label Generator")
st.info("Upload your Excel file (Бочки) to generate labels.")
>>>>>>> 32fc247 (Fixed pallet weights and added ffill)

# Custom CSS to center elements and style buttons
st.markdown("""
    <style>
    /* Center the main title */
    .stTitle {
        text-align: center;
    }
    
    /* Center the file uploader label */
    .stFileUploader label {
        display: flex;
        justify-content: center;
        font-weight: bold;
        font-size: 1.2rem;
    }

    /* Style the download buttons */
    .stDownloadButton > button {
        width: 100%;
        height: 3.5em;
        background-color: #2e7d32;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        border: none;
    }
    
    /* Style the generate buttons */
    .stButton > button {
        width: 100%;
        border-radius: 8px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📦 Erebuni Label Generator")
st.write("---")

# Main Area Upload (Moved from Sidebar to Center)
uploaded_file = st.file_uploader(
    "Upload Center: Choose Excel File (Бочки)", 
    type=['xlsx'],
    help="Limit 200MB per file • XLSX"
)

if uploaded_file:
    with st.status("Reading data and applying weights...", expanded=False) as status:
        df = pd.read_excel(uploaded_file, skiprows=4)
        df.columns = [" ".join(str(c).split()) for c in df.columns]
        
        for col in ['Нетто соуса на паллете', 'Брутто паллета']:
            if col in df.columns:
                df[col] = df[col].ffill()
                
        df = df[df['Номер Партии'].notna()].copy()
        df = df.reset_index(drop=True)
        status.update(label="✅ Data Processed Successfully!", state="complete")

    m1, m2, m3 = st.columns(3)
    m1.metric("Labels", len(df))
    m2.metric("Pallets", len(df) // 4)
    if m3.button("🗑️ Restart"):
        st.rerun()

    with st.expander("📄 View Data Table Preview"):
        st.dataframe(df, use_container_width=True)

    st.markdown("### 📥 Download Generated Files")
    
    # df = pd.read_excel(uploaded_file, skiprows=4)
    # df.columns = [" ".join(str(c).split()) for c in df.columns]
    # for col in ['Нетто соуса на паллете', 'Брутто паллета']:
    #     if col in df.columns:
    #         df[col] = df[col].ffill()
    # df = df[df['Номер Партии'].notna()].copy()
    # df = df.reset_index(drop=True)
    # st.success(f"Loaded {len(df)} labels from Excel.")
    
    col1, col2 = st.columns(2)
    
    
    with col1:
        if st.button("🛠️ Prepare PDF Labels"):
            pdf_data = create_pdf_file(df)
            st.download_button(
                label="💾 Download PDF",
                data=pdf_data,
                file_name="Erebuni_Labels.pdf",
                mime="application/pdf"
            )
            
    with col2:
        if st.button("🛠️ Prepare Word Labels"):
            word_data = create_word_file(df)
            st.download_button(
                label="💾 Download Word",
                data=word_data,
                file_name="Erebuni_Labels.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    # This shows when the app is empty
    st.info("Please upload your Excel file above to generate the labels and pallet sheets.")
