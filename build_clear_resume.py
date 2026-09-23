"""Build the conservative resume and its plain-text equivalent.

Sources: existing resume .tex files and Ford-Responsibilities-Summary.tex.
The user confirmed Ford's current Data Scientist role, CereVu's side-contract
arrangement and titles, Computer Engineering master's (May 2021), and Ph.D.
graduation (December 2023). The university's May 2021 commencement program
identifies the master's degree as Master of Engineering.
"""
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output" / "resume"
OUT.mkdir(parents=True, exist_ok=True)
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
sec.top_margin = sec.bottom_margin = Inches(0.62)
sec.left_margin = sec.right_margin = Inches(0.7)
normal = doc.styles['Normal']
normal.font.name = 'Arial'
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor(0, 0, 0)
normal.paragraph_format.space_after = Pt(4)
normal.paragraph_format.line_spacing = 1.06
for name, size in [('Title', 22), ('Heading 1', 11.5), ('Heading 2', 10.5)]:
    s = doc.styles[name]
    s.font.name = 'Arial'
    s.font.size = Pt(size)
    s.font.bold = True
    s.font.color.rgb = RGBColor(0, 0, 0)
    s.paragraph_format.space_before = Pt(11 if name == 'Heading 1' else 5)
    s.paragraph_format.space_after = Pt(4)
    s.paragraph_format.keep_with_next = True
doc.styles['Title'].paragraph_format.space_before = Pt(0)
# The bundled Word template carries a border in its Title style.
for style in doc.styles:
    for border in list(style.element.iter(qn('w:pBdr'))):
        border.getparent().remove(border)
for level in doc.part.numbering_part.element.iter(qn('w:lvl')):
    fmt = level.find(qn('w:numFmt'))
    if fmt is not None and fmt.get(qn('w:val')) == 'bullet':
        level.find(qn('w:lvlText')).set(qn('w:val'), '\u2022')
        props = level.find(qn('w:rPr'))
        if props is not None:
            fonts = props.find(qn('w:rFonts'))
            if fonts is not None:
                fonts.set(qn('w:ascii'), 'Arial')
                fonts.set(qn('w:hAnsi'), 'Arial')
doc.core_properties.author = 'Moein Razavi'
doc.core_properties.title = 'Moein Razavi Resume'
doc.core_properties.subject = 'Data Science and Software Development'
doc.core_properties.keywords = ''
plain = []

def para(text, style=None, bold_prefix=None):
    p = doc.add_paragraph(style=style)
    if bold_prefix:
        p.add_run(bold_prefix).bold = True
        p.add_run(text[len(bold_prefix):])
    else:
        p.add_run(text)
    plain.append(text)
    return p

def heading(text):
    plain.append('')
    return para(text, 'Heading 1')

def role(company, title, dates):
    para(company, 'Heading 2')
    p = para(f'{title} | {dates}')
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_after = Pt(5)

def bullet(text):
    p = doc.add_paragraph(style='List Bullet')
    p.add_run(text)
    p.paragraph_format.left_indent = Inches(0.15)
    p.paragraph_format.first_line_indent = Inches(-0.15)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_together = True
    plain.append('- ' + text)

para('Moein Razavi', 'Title')
para('McKinney, TX | (979) 676-7486 | razavi.moein94@gmail.com')
para('linkedin.com/in/moein-razavi | github.com/moeinrazavi')

heading('Professional Summary')
para('Data scientist and software developer with a Ph.D. in Industrial Engineering. Experience building machine learning models, data pipelines, backend services, and iOS applications. Work spans electric vehicle analytics at Ford, wearable sensor applications at CereVu Medical, and university research in stress detection.')

heading('Technical Skills')
for label, content in [
    ('Programming', 'Python, SQL, Swift, TypeScript'),
    ('Machine Learning', 'PyTorch, TensorFlow, scikit-learn, XGBoost, time-series forecasting, Optuna'),
    ('AI Applications', 'LangChain, retrieval-augmented generation (RAG), text-to-SQL, FAISS'),
    ('Data and Cloud', 'Google Cloud Platform (GCP), BigQuery, Vertex AI, Cloud Run, PySpark, Apache Airflow'),
    ('Software Development', 'Flask, FastAPI, REST APIs, WebSocket, SwiftUI, React, Docker, Tekton, GitHub Actions'),
]:
    para(f'{label}: {content}', bold_prefix=label + ':')

heading('Professional Experience')
role('Ford Motor Company', 'Data Scientist', 'Aug 2023 - Present')
bullet('Developed time-series forecasting models for electric vehicle (EV) analytics using transformer architectures, Hugging Face, and Optuna for hyperparameter tuning.')
bullet('Built data pipelines with PySpark, BigQuery, and Airflow to process vehicle telemetry and prepare inputs and outputs for energy simulations.')
bullet('Developed an EV analytics chatbot that routed questions to SQL queries or document retrieval using LangChain and retrieval-augmented generation.')
bullet('Built a utility rate recommendation application with Flask APIs, BigQuery workflows, and a Dash interface; deployed the application on Cloud Run using Tekton pipelines.')
bullet('Created Looker Studio dashboards to report EV discharge activity and developed REST APIs for operational dashboards.')

role('CereVu Medical, Inc.', 'Software Developer / Research Engineer (Contract)', 'Jun 2023 - Sep 2024')
bullet('Developed Vitality and Clinical iOS applications in SwiftUI to display physiological signals from wearable sensors, with CoreBluetooth integration and data visualization.')
bullet('Migrated the Clinical application from UIKit to SwiftUI and implemented role-based screens, data export, and calibration features.')
bullet('Built Python services with Flask, FastAPI, and WebSocket for biosignal streaming, and developed a React/Next.js dashboard for reviewing patient data.')

p = heading('Research Experience')
p.paragraph_format.page_break_before = True
role('Texas A&M University', 'Research Assistant', 'Sep 2018 - May 2024')
bullet('Developed machine learning workflows for stress detection from wearable heart-rate and motion data, including signal preprocessing, feature extraction, and model evaluation.')
bullet('Co-developed OpenSync, an open-source platform for synchronizing EEG, eye-tracking, physiological signals, and experiment events in neuroscience studies.')
bullet('Developed computer vision methods for detecting face masks and estimating physical distance in video as part of COVID-19 monitoring research.')
bullet('Co-authored a scoping review of machine learning, deep learning, and data preprocessing methods for stress and stress-related mental disorder research.')

heading('Selected Projects')
para('Inspecalytics', 'Heading 2')
bullet('Developed a SwiftUI field-inspection application with photo annotation, offline data storage, and automated Word and PowerPoint report generation.')
bullet('Configured Fastlane and GitHub Actions to build and distribute TestFlight releases.')
para('Clink Social', 'Heading 2')
bullet('Developed a SwiftUI social application with Firebase authentication, Firestore data storage, messaging, and location-based discovery.')
bullet('Implemented TypeScript cloud functions for feed aggregation and push notifications.')

heading('Education')
para('Doctor of Philosophy (Ph.D.), Industrial Engineering', 'Heading 2')
para('Texas A&M University | Dec 2023')
para('Master of Engineering, Computer Engineering', 'Heading 2')
para('Texas A&M University | May 2021')

heading('Selected Publications')
para('Razavi, M., et al. (2024). Machine Learning, Deep Learning, and Data Preprocessing Techniques for Detecting, Predicting, and Monitoring Stress and Stress-Related Mental Disorders: Scoping Review. JMIR Mental Health, 11, e53714. doi:10.2196/53714.')
para('Razavi, M., et al. (2022). OpenSync: An open-source platform for synchronizing multiple measures in neuroscience experiments. Journal of Neuroscience Methods, 369, 109458. doi:10.1016/j.jneumeth.2021.109458.')

# Single-column body text; no tables, graphics, contact headers, or text boxes.
assert len(doc.tables) == 0
doc.save(OUT / 'Moein-Razavi-Clear-Resume.docx')
(OUT / 'Moein-Razavi-Clear-Resume.txt').write_text('\n'.join(plain) + '\n')
print(f'Created resume in {OUT}; {len(" ".join(plain).split())} words')
