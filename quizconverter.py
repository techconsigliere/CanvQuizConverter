import io
import re
import xml.etree.ElementTree as ET
import zipfile
import streamlit as st
from docx import Document

def parse_docx(uploaded_file):
    doc = Document(uploaded_file)
    questions = []
    curr = None

    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            continue

        q_m = re.match(r'^(\d+)\.\s*(.*)', txt)
        if q_m:
            if curr:
                questions.append(curr)
            curr = {
                'num': q_m.group(1), 
                'text': q_m.group(2), 
                'choices': [], 
                'correct_letter': None
            }
            continue

        c_m = re.match(r'^(\*?)([a-dA-D])\.\s*(.*)', txt)
        if c_m and curr:
            is_ans = bool(c_m.group(1))
            cl = c_m.group(2).lower()
            ct = c_m.group(3)
            qnum = curr['num']
            cid = f"choice_{qnum}_{cl}"
            
            c_dict = {
                'id': cid, 
                'letter': cl, 
                'text': ct, 
                'is_correct': is_ans
            }
            curr['choices'].append(c_dict)
            if is_ans:
                curr['correct_letter'] = cl

    if curr:
        questions.append(curr)
    return questions

def create_qti_xml(questions, quiz_title):
    ns_qti = "http://www.imsglobal.org/xsd/ims_qtiasiv1p2"
    ns_xsi = "http://www.w3.org/2001/XMLSchema-instance"
    
    attr = {'xmlns': ns_qti, 'xmlns:xsi': ns_xsi}
    qti = ET.Element('questestinterop', attr)
    
    as_attr = {'ident': 'canvas_export', 'title': quiz_title}
    assessment = ET.SubElement(qti, 'assessment', as_attr)
    
    sec_attr = {'ident': 'root', 'title': 'Main Section'}
    section = ET.SubElement(assessment, 'section', sec_attr)

    for q in questions:
        qnum = q['num']
        it_attr = {'ident': f"q_{qnum}", 'title': f"Q {qnum}"}
        item = ET.SubElement(section, 'item', it_attr)

        correct_ids = [c['id'] for c in q['choices'] if c['is_correct']]
        is_multi = len(correct_ids) > 1

        meta = ET.SubElement(item, 'itemmetadata')
        qmeta = ET.SubElement(meta, 'qtimetadata')
        
        m_fields = [
            ('question_type', 'multiple_answers_question' if is_multi else 'multiple_choice_question'),
            ('points_possible', '1.0')
        ]
        for label, entry in m_fields:
            field = ET.SubElement(qmeta, 'qtimetadatafield')
            ET.SubElement(field, 'fieldlabel').text = label
            ET.SubElement(field, 'fieldentry').text = entry

        pres = ET.SubElement(item, 'presentation', {'label': 'color'})
        mat = ET.SubElement(pres, 'material')
        m_txt = ET.SubElement(mat, 'mattext', {'texttype': 'text/plain'})
        m_txt.text = q['text']

        rlid_attr = {'ident': 'response1', 'rcardinality': 'Multiple' if is_multi else 'Single'}
        rlid = ET.SubElement(pres, 'response_lid', rlid_attr)
        render = ET.SubElement(rlid, 'render_choice')

        for choice in q['choices']:
            ch_id = choice['id']
            ch_txt = choice['text']
            lbl_attr = {'ident': ch_id}
            r_lbl = ET.SubElement(render, 'response_label', lbl_attr)
            c_mat = ET.SubElement(r_lbl, 'material')
            c_mtxt = ET.SubElement(c_mat, 'mattext', {'texttype': 'text/plain'})
            c_mtxt.text = ch_txt

        resproc = ET.SubElement(item, 'resprocessing')
        outcomes = ET.SubElement(resproc, 'outcomes')
        dv_attr = {'vartype': 'Decimal', 'defaultval': '0.0'}
        ET.SubElement(outcomes, 'decvar', dv_attr)
        
        rcond = ET.SubElement(resproc, 'respcondition', {'continue': 'No'})
        cvar = ET.SubElement(rcond, 'conditionvar')

        if is_multi:
            # Multiple correct answers: student must match the full set, no more, no less.
            and_el = ET.SubElement(cvar, 'and')
            for correct_id in correct_ids:
                ET.SubElement(and_el, 'varequal', {'respident': 'response1'}).text = correct_id
            for choice in q['choices']:
                if not choice['is_correct']:
                    not_el = ET.SubElement(and_el, 'not')
                    ET.SubElement(not_el, 'varequal', {'respident': 'response1'}).text = choice['id']
        elif correct_ids:
            ET.SubElement(cvar, 'varequal', {'respident': 'response1'}).text = correct_ids[0]

        ET.SubElement(rcond, 'setvar', {'action': 'Set', 'varname': 'SCORE'}).text = '100'

    return ET.tostring(qti, encoding='utf-8', method='xml')

def create_manifest_xml():
    ns_imscp = "http://www.imsglobal.org/xsd/imsccv1p1/imscp_v1p1"
    ns_xsi = "http://www.w3.org/2001/XMLSchema-instance"
    m_attr = {'identifier': 'canvas_manifest', 'xmlns': ns_imscp, 'xmlns:xsi': ns_xsi}
    manifest = ET.Element('manifest', m_attr)
    resources = ET.SubElement(manifest, 'resources')
    r_attr = {'identifier': 'quiz_res', 'type': 'imsqti_xmlv1p2', 'href': 'quiz.xml'}
    ET.SubElement(resources, 'resource', r_attr)
    return ET.tostring(manifest, encoding='utf-8', method='xml')

# --- STREAMLIT WEB INTERFACE ---
st.set_page_config(page_title="Canvas Quiz Converter", page_icon="📝", layout="centered")
st.title("📝 Canvas Quiz Converter")

exp_p1 = "Upload a .docx file, preview questions and correct answers, "
exp_p2 = "and convert it into a Canvas QTI package that can easily be imported into your Canvas course."
st.write(f"{exp_p1}{exp_p2}")

uploaded_file = st.file_uploader("Choose your quiz Word document", type=["docx"])

if uploaded_file is not None:
    questions = parse_docx(uploaded_file)
    quiz_name = uploaded_file.name.replace(".docx", "")
    
    if len(questions) == 0:
        st.error("❌ No questions detected.")
    else:
        st.success(f"✅ Loaded: {uploaded_file.name}")
        st.subheader("👀 Quiz Preview")
        
        for q in questions:
            qnum = q['num']
            qtxt = q['text']
            title_text = f"Question {qnum}: {qtxt[:40]}{'...' if len(qtxt) > 40 else ''}"
            with st.expander(title_text, expanded=True):
                st.markdown(f"**Question {qnum}:** {qtxt}")
                
                for choice in q['choices']:
                    let = choice['letter']
                    txt = choice['text']
                    is_ans = choice['is_correct']
                    if is_ans:
                        st.markdown(f"✅ :green[**{let}.** {txt}]")
                    else:
                        st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**{let}.** {txt}")
                if not q['correct_letter']:
                    st.error("⚠️ Warning: No correct answer marked!")
                elif len(q['choices']) and sum(1 for c in q['choices'] if c['is_correct']) > 1:
                    st.warning("⚠️ Multiple correct answers marked. This will export as a multiple-answer question (student must select all marked choices).")

        st.divider()
        st.subheader("💾 Export to Canvas")
        
        lbl = "Please provide a description for your exported quiz."
        custom_filename = st.text_input(label=lbl, value=f"{quiz_name}_canvas_quiz")
        
        is_zip = custom_filename.endswith(".zip")
        final_filename = custom_filename if is_zip else f"{custom_filename}.zip"
        
        zip_buffer = io.BytesIO()
        qti_xml_data = create_qti_xml(questions, quiz_name)
        manifest_data = create_manifest_xml()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as qti_zip:
            qti_zip.writestr('quiz.xml', qti_xml_data)
            qti_zip.writestr('imsmanifest.xml', manifest_data)
            
        st.download_button(
            label="🚀 Convert & Download Canvas QTI File", 
            data=zip_buffer.getvalue(), 
            file_name=final_filename, 
            mime="application/zip", 
            use_container_width=True
        )