from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import google.generativeai as genai
import os
import re
from database import Recommendation_English , Recommendation_Arabic

os.environ['API_KEY'] = 'AIzaSyCVVe2FwYmaaDG61RAQ-e8pOvIs8CzsrME'
genai.configure(api_key=os.environ['API_KEY'])

try:
    model = genai.GenerativeModel('gemini-1.5-pro')
except Exception as e:
    print(f"Error initializing model: {e}")
    model = None

def generate_recommendations_english(data):
        if model is None:
            return "Unable to generate recommendations due to model initialization error."

        prompt = f"""
        Based on the following energy audit data, provide 5-7 specific recommendations for saving power and improving energy efficiency:

            Accommodation: {data.get('type_of_accommodation')}
            number of residents: {data['number_of_residents']}
            Year of Construction: {data['year_of_construction']}
            Number of Bedrooms: {data['number_of_bedrooms']}
            Number of Floors: {data['number_of_floors']}
            Outdoor Garden: {data['outdoor_garden']}
            Swimming Pool: {data['swimming_pool']}
            Air Conditioning Systems: {data['ac_systems']}
            Lighting: {data['lighting']}
            Water Taps: {data['water_taps']}
            Water Heaters: {data['water_heaters']}
            Other Notes: {data['other']}'

        Please provide actionable and specific recommendations, benefits andimplementation. Format the recommendations as a bullet point list in the following format.
        AC-System.:
        Lighting:
        
        Water Taps:

        Water Heaters:

        Other Observations:
        
        Do not write anything before this and dont bold any sentences/words with **. Use - to bullet and dont use brackets anywhere
        """

        response = model.generate_content(prompt)
        cleaned_response = response.text.replace("*", "")
        cleaned_response = re.sub(r'\[.*?\]', '', cleaned_response) 
        return cleaned_response

def set_cell_shading( cell, color):
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), color)
        cell._element.get_or_add_tcPr().append(shading)

def set_ltr( paragraph):
        p = paragraph._element
        pPr = p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '0')
        pPr.append(bidi)
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

def set_rtl( paragraph):
        p = paragraph._element
        pPr = p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

def set_borders( cell):
        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tcBorders.append(border)
        tcPr.append(tcBorders)

def set_paragraph_spacing( paragraph):
        pPr = paragraph._element.get_or_add_pPr()
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:line'), '360')  # 1.5 * 240 TWIPS
        pPr.append(spacing)

def add_hyperlink(paragraph, url, text):
    # Create the w:hyperlink tag and add required attributes
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Add color
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')  # Hex value for blue color
    rPr.append(color)

    # Add underline
    underline = OxmlElement('w:u')
    underline.set(qn('w:val'), 'single')
    rPr.append(underline)

    new_run.append(rPr)

    text_element = OxmlElement('w:t')
    text_element.text = text
    new_run.append(text_element)
    hyperlink.append(new_run)

    # Append the hyperlink to the paragraph
    paragraph._element.append(hyperlink)
    return paragraph


def create_report_english(data, recommendations):
        doc = Document()

        # Add logos to the header
        section = doc.sections[0]
        header = section.header

        logo_paragraph = header.paragraphs[0]
        logo_run = logo_paragraph.add_run()
        logo_run.add_picture(r"./rak.png", width=Inches(1.6))
        logo_run.add_text(" " * 70)  # Adjust the number of spaces to control the gap
        logo_run.add_picture(r"./mun.png", width=Inches(2.0))
        logo_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Title
        title = doc.add_heading(('Manzili Energy Audit Service Report'), level=1)
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Report Number
        report_number_paragraph = doc.add_paragraph(f"{('Report Number')}: {data['report_number']}")
        report_number_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Overview
        overview_heading = doc.add_heading(('Overview'), level=2)
        overview_paragraph = doc.add_paragraph((
            "This report summarizes the results and recommendations following the energy audit conducted in your home as part of the Manzili home energy consultancy service in Ras Al Khaimah. The goal of the audit is to help reduce your electricity and water bills and make your home more comfortable and modern."
        ))

        # Audit Details
        audit_details_heading = doc.add_heading(('Audit Details'), level=2)
        audit_details_heading.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        audit_table = doc.add_table(rows=1, cols=4)
        hdr_cells = audit_table.rows[0].cells
        hdr_cells[0].text = ('Item')
        hdr_cells[1].text = ('Details')
        hdr_cells[2].text = ('Item')
        hdr_cells[3].text = ('Details')

        for cell in hdr_cells:
            set_cell_shading(cell, "D3D3D3")  # Set header cell color to grey
            for paragraph in cell.paragraphs:
                set_ltr(paragraph)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            set_borders(cell)

        audit_fields = [
            ('date_of_audit', 'report_number'),
            ('homeowner', 'contact_number'),
            ('location', 'type_of_accommodation'),
            ('house_number', 'year_of_construction'),
            ('number_of_bedrooms', 'number_of_floors')
        ]

        for field_pair in audit_fields:
            row_cells = audit_table.add_row().cells
            row_cells[0].text = (field_pair[0].replace('_', ' ').title())
            row_cells[1].text = data[field_pair[0]]
            row_cells[2].text = (field_pair[1].replace('_', ' ').title())
            row_cells[3].text = data[field_pair[1]]
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    set_ltr(paragraph)
                    set_paragraph_spacing(paragraph)
                set_borders(cell)
            set_cell_shading(row_cells[0], "D3D3D3")
            set_cell_shading(row_cells[2], "D3D3D3")

        # Notes
        notes_heading = doc.add_heading(('Notes'), level=2)
        notes_table = doc.add_table(rows=1, cols=2)
        hdr_cells = notes_table.rows[0].cells
        hdr_cells[0].text = ('Item')
        hdr_cells[1].text = ('Details')

        for cell in hdr_cells:
            set_cell_shading(cell, "D3D3D3")  # Set header cell color to grey
            for paragraph in cell.paragraphs:
                set_ltr(paragraph)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            set_borders(cell)

        for key in ['outdoor_garden', 'swimming_pool', 'ac_systems', 'lighting', 'water_taps', 'water_heaters']:
            row_cells = notes_table.add_row().cells
            row_cells[0].text = (key.replace('_', ' ').title())
            row_cells[1].text = data[key]
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    set_ltr(paragraph)
                    set_paragraph_spacing(paragraph)
                set_borders(cell)
            set_cell_shading(row_cells[0], "D3D3D3")

        # Recommendations
        i = 0
        recommendations_heading = doc.add_heading(('Recommendations'), level=2)
        recommendations_table = doc.add_table(rows=1, cols=3)
        hdr_cells = recommendations_table.rows[0].cells
        hdr_cells[0].text = ('Recommendations')
        hdr_cells[1].text = ('Benefits')
        hdr_cells[2].text = ('Implementation')
        for cell in hdr_cells:
            set_cell_shading(cell, "D3D3D3")  # Set header cell color to grey
            for paragraph in cell.paragraphs:
                set_ltr(paragraph)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            set_borders(cell)
        i+=1
        
        priority_list = ['High Priority', 'Medium Priority', 'Low Priority']
        for priority in priority_list:
            row_cells = recommendations_table.add_row().cells
            recommendations_table.cell(i, 0).merge(recommendations_table.cell(i, 2))
            row_cells[0].text = priority
            set_paragraph_spacing(paragraph)
            set_cell_shading(row_cells[0], "D3D3D3")
            for cell in row_cells:
                set_cell_shading(cell, "D3D3D3")  # Set header cell color to grey
                for paragraph in cell.paragraphs:
                    set_ltr(paragraph)
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                set_borders(cell)
            i+=1

            for key, value in Recommendation_English.items():
                if key in data:
                    if data[f'dropdown_{key}'] == priority:
                        string = fr'{value[0]} in {data[f"input_{key}"]}'
                        row_cells = recommendations_table.add_row().cells
                        row_cells[0].text = (string)
                        row_cells[1].text = (value[1])
                        if len(Recommendation_English[key])==4:
                            hyp_text = value[2].split('\n')
                            row_cells[2].text = (hyp_text[0])
                            add_hyperlink(row_cells[2].paragraphs[0],value[3], hyp_text[1])
                        else:
                            row_cells[2].text = (value[2])
                        for cell in row_cells:
                            set_borders(cell)
                        i+=1 
             


        # AI-Generated Recommendations
        ai_recommendations_heading = doc.add_heading(('AI-Generated Recommendations'), level=2)
        ai_recommendations = doc.add_paragraph(recommendations)
        set_ltr(ai_recommendations)
        set_paragraph_spacing(ai_recommendations)

        # Disclaimer
        disclaimer_heading = doc.add_heading(('Disclaimer'), level=2)
        disclaimer_paragraph_1 = doc.add_paragraph((
            "This report is based on visual observations of the main equipment related to energy and water in your home by the Ras Al Khaimah Municipality. The observations do not include any detailed measurements or analyses."
        ))
        disclaimer_paragraph_2 = doc.add_paragraph((
            "Potential savings indicated in the report are estimates and not guaranteed. There is no obligation to implement any recommendations, and the Ras Al Khaimah Municipality will not be liable for any actions taken by the homeowner or any other party."
        ))
        disclaimer_paragraph_3 = doc.add_paragraph((
            "The information provided is based on available data from the Ras Al Khaimah Municipality and recommended suppliers and contractors. The municipality welcomes feedback on the listed companies and suggestions for new companies to be added to the list. For any suggestions, please email manzily@mun.rak.ae."
        ))

        # Save the document
        filename = f'Manzili_Energy_Audit_Report_{data["report_number"]}.docx'
        filepath = os.path.join(os.path.dirname(__file__), filename)

        if os.path.exists(filepath):
            os.remove(filepath)

        doc.save(filepath)


'''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                                                      ARABIC
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////'''



def generate_recommendations_arabic( data):
        if model is None:
            return "تعذر إنشاء التوصيات بسبب خطأ في تهيئة النموذج."

        prompt = f"""
        استنادًا إلى بيانات تدقيق الطاقة التالية، قدم 5-7 توصيات محددة لتوفير الطاقة وتحسين كفاءة الطاقة:

        نوع الإقامة: {data['نوع_الإقامة']}
        عدد السكان: {data['عدد_السكان']}
        سنة البناء: {data['سنة_البناء']}
        عدد غرف النوم: {data['عدد_غرف_النوم']}
        عدد الطوابق: {data['عدد_الطوابق']}
        حديقة خارجية: {data['حديقة_خارجية']}
        حمام سباحة: {data['حمام_سباحة']}
        أنظمة تكييف الهواء: {data['أنظمة_تكييف']}
        الإضاءة: {data['إضاءة']}
        حنفيات المياه: {data['حنفيات_المياه']}
        سخانات المياه: {data['سخانات_المياه']}
        ملاحظات أخرى: {data['أخرى']}
        يرجى تقديم توصيات محددة وقابلة للتنفيذ، الفوائد والتنفيذ. قم بتنسيق التوصيات على شكل قائمة نقطية في الشكل التالي.
        
        نظام التكييف:
        
        الإضاءة:
        
        الحنفيات:
        
        سخانات المياه:
        
        ملاحظات أخرى:
        
        لا تكتب أي شيء قبل هذا ولا تجعل أي جمل/كلمات غامقة. استخدم - للتنقيط ولا تستخدم أقواس في أي مكان    
        """

        response = model.generate_content(prompt)
        return response.text

def create_report_arabic( data, recommendations):
        doc = Document()

        # Add logos to the header
        section = doc.sections[0]
        header = section.header

        logo_paragraph = header.paragraphs[0]
        logo_run = logo_paragraph.add_run()
        logo_run.add_picture(r"./rak.png", width=Inches(1.6))
        logo_run.add_text(" " * 70)  # Adjust the number of spaces to control the gap
        logo_run.add_picture(r"./mun.png", width=Inches(2.0))
        logo_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Title
        title = doc.add_heading('تقرير منزلي لخدمة تدقيق الطاقة منزلية ', level=1)
        set_rtl(title)  # Set RTL for title
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Report Number
        report_number_paragraph = doc.add_paragraph(f"رقم التقرير :{data['رقم_التقرير']}")
        set_rtl(report_number_paragraph)  # Set RTL for report number
        report_number_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Overview
        overview_heading = doc.add_heading('نظرة عامة', level=2)
        set_rtl(overview_heading)
        overview_paragraph = doc.add_paragraph(
            "يُلخص هذا التقرير النتائج والتوصيات بعد تدقيق الطاقة الذي أُجري في منزلك كجزء من خدمة استشارات طاقة منزلي في رأس الخيمة. الهدف من التدقيق هو المساعدة في تقليل فواتير الكهرباء والمياه وجعل منزلك أكثر راحة وحداثة."
        )
        #set_rtl(overview_paragraph)
        set_paragraph_spacing(overview_paragraph)

        # Audit Details
        audit_details_heading = doc.add_heading('تفاصيل التدقيق', level=2)
        audit_details_heading.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        set_rtl(audit_details_heading)
        audit_table = doc.add_table(rows=1, cols=4)
        hdr_cells = audit_table.rows[0].cells
        hdr_cells[0].text = 'التفاصيل'
        hdr_cells[1].text = 'العنصر'
        hdr_cells[2].text = 'التفاصيل'
        hdr_cells[3].text = 'العنصر'

        for cell in hdr_cells:
            set_cell_shading(cell, "D3D3D3")  # Set header cell color to grey
            for paragraph in cell.paragraphs:
                set_rtl(paragraph)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            set_borders(cell)

        audit_fields = [
            ('تاريخ_التدقيق', 'رقم_التقرير'),
            ('صاحب_المنزل', 'رقم_الاتصال'),
            ('الموقع', 'نوع_الإقامة'),
            ('رقم_المنزل', 'سنة_البناء'),
            ('عدد_غرف_النوم', 'عدد_الطوابق')
        ]

        for field_pair in audit_fields:
            row_cells = audit_table.add_row().cells
            row_cells[1].text = field_pair[0].replace('_', ' ').title()
            row_cells[0].text = data[field_pair[0]]
            row_cells[3].text = field_pair[1].replace('_', ' ').title()
            row_cells[2].text = data[field_pair[1]]
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    set_rtl(paragraph)
                    set_paragraph_spacing(paragraph)
                set_borders(cell)
            set_cell_shading(row_cells[1], "D3D3D3")
            set_cell_shading(row_cells[3], "D3D3D3")

        # Notes
        notes_heading = doc.add_heading('الملاحظات', level=2)
        set_rtl(notes_heading)
        notes_table = doc.add_table(rows=1, cols=2)
        hdr_cells = notes_table.rows[0].cells
        hdr_cells[1].text = 'العنصر'
        hdr_cells[0].text = 'التفاصيل'

        for cell in hdr_cells:
            set_cell_shading(cell, "D3D3D3")  # Set header cell color to grey
            for paragraph in cell.paragraphs:
                set_rtl(paragraph)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            set_borders(cell)

        for key in ['حديقة_خارجية', 'حمام_سباحة', 'أنظمة_تكييف', 'إضاءة', 'حنفيات_المياه', 'سخانات_المياه']:
            row_cells = notes_table.add_row().cells
            row_cells[1].text = key.replace('_', ' ').title()
            row_cells[0].text = data[key]
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    set_rtl(paragraph)
                    set_paragraph_spacing(paragraph)
                set_borders(cell)
            set_cell_shading(row_cells[1], "D3D3D3")

        # Recommendations
        i = 0
        recommendations_heading = doc.add_heading((' التوصية'), level=2)
        recommendations_table = doc.add_table(rows=1, cols=3)
        hdr_cells = recommendations_table.rows[0].cells
        hdr_cells[0].text = ('التنفيذ')
        hdr_cells[1].text = ('الفوائد')
        hdr_cells[2].text = (' التوصية')
        for cell in hdr_cells:
            set_cell_shading(cell, "D3D3D3")  # Set header cell color to grey
            for paragraph in cell.paragraphs:
                set_ltr(paragraph)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            set_borders(cell)
        i+=1
        
        priority_list = ['أولوية قصوى', 'أولوية متوسطة', 'أولوية منخفضة']
        for priority in priority_list:
            row_cells = recommendations_table.add_row().cells
            recommendations_table.cell(i, 0).merge(recommendations_table.cell(i, 2))
            row_cells[0].text = priority
            set_paragraph_spacing(paragraph)
            set_cell_shading(row_cells[0], "D3D3D3")
            for cell in row_cells:
                set_cell_shading(cell, "D3D3D3")  # Set header cell color to grey
                for paragraph in cell.paragraphs:
                    set_ltr(paragraph)
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                set_borders(cell)
            i+=1

            for key, value in Recommendation_Arabic.items():
                if key in data:
                    if data[f'dropdown_{key}'] == priority:
                        string = fr'{value[0]} في {data[f"input_{key}"]}'
                        row_cells = recommendations_table.add_row().cells
                        row_cells[0].text = (string)
                        row_cells[1].text = (value[1])
                        if len(Recommendation_Arabic[key])==4:
                            hyp_text = value[2].split('\n')
                            row_cells[2].text = (hyp_text[0])
                            add_hyperlink(row_cells[2].paragraphs[0],value[3], hyp_text[1])
                        else:
                            row_cells[2].text = (value[2])
                        for cell in row_cells:
                            set_borders(cell)
                        i+=1 
                        
                        
        # AI-Generated Recommendations
        ai_recommendations_heading = doc.add_heading('التوصيات المولّدة بواسطة الذكاء الاصطناعي', level=2)
        set_rtl(ai_recommendations_heading)
        ai_recommendations = doc.add_paragraph(recommendations)
        set_rtl(ai_recommendations)
        set_paragraph_spacing(ai_recommendations)

        # Disclaimer
        disclaimer_heading = doc.add_heading('تنويه', level=2)
        set_rtl(disclaimer_heading)
        disclaimer_paragraph_1 = doc.add_paragraph(
            "يستند هذا التقرير إلى الملاحظات البصرية للمعدات الرئيسية المتعلقة بالطاقة والمياه في منزلك من قبل بلدية رأس الخيمة. لا تشمل الملاحظات أي قياسات أو تحاليل مفصلة."
        )
        set_rtl(disclaimer_paragraph_1)
        set_paragraph_spacing(disclaimer_paragraph_1)
        disclaimer_paragraph_2 = doc.add_paragraph(
            "المدخرات المحتملة المشار إليها في التقرير هي تقديرات وليست مضمونة. لا يوجد التزام بتنفيذ أي توصيات، ولن تكون بلدية رأس الخيمة مسؤولة عن أي إجراءات يتخذها صاحب المنزل أو أي طرف آخر."
        )
        set_rtl(disclaimer_paragraph_2)
        set_paragraph_spacing(disclaimer_paragraph_2)
        disclaimer_paragraph_3 = doc.add_paragraph(
            "المعلومات المقدمة تستند إلى البيانات المتاحة من بلدية رأس الخيمة والموردين والمقاولين الموصى بهم. ترحب البلدية بالتعليقات حول الشركات المدرجة والاقتراحات لإضافة شركات جديدة إلى القائمة. لأي اقتراحات، يرجى إرسال بريد إلكتروني إلى manzily@mun.rak.ae."
        )
        set_rtl(disclaimer_paragraph_3)
        set_paragraph_spacing(disclaimer_paragraph_3)

        # Save the document
        filename = f'Manzili_Energy_Audit_Report_{data["رقم_التقرير"]}.docx'
        filepath = os.path.join(os.path.dirname(__file__), filename)

        if os.path.exists(filepath):
            os.remove(filepath)

        doc.save(filepath)
