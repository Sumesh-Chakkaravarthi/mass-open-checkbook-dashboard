#!/usr/bin/env python3
"""
Generate professional cover letters as PDFs for Sumesh Chakkaravarthi Purushothaman.
Creates 5 role-specific cover letters + 1 generic template.
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.colors import HexColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
from datetime import datetime

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Cover_Letters")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Color palette
DARK_NAVY = HexColor("#1a1a2e")
ACCENT_BLUE = HexColor("#2d5f8a")
MEDIUM_GRAY = HexColor("#444444")
LIGHT_GRAY = HexColor("#888888")
LINE_COLOR = HexColor("#2d5f8a")

# Date
current_date = datetime.now().strftime("%B %d, %Y")


def get_styles():
    """Create professional paragraph styles."""
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name='ContactName',
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        textColor=DARK_NAVY,
        alignment=TA_CENTER,
        spaceAfter=2,
    ))

    styles.add(ParagraphStyle(
        name='ContactInfo',
        fontName='Helvetica',
        fontSize=8.5,
        leading=12,
        textColor=ACCENT_BLUE,
        alignment=TA_CENTER,
        spaceAfter=2,
    ))

    styles.add(ParagraphStyle(
        name='DateLine',
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=MEDIUM_GRAY,
        alignment=TA_LEFT,
        spaceAfter=6,
        spaceBefore=8,
    ))

    styles.add(ParagraphStyle(
        name='Salutation',
        fontName='Helvetica-Bold',
        fontSize=10.5,
        leading=14,
        textColor=DARK_NAVY,
        alignment=TA_LEFT,
        spaceAfter=8,
    ))

    styles.add(ParagraphStyle(
        name='BodyText_Custom',
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=MEDIUM_GRAY,
        alignment=TA_JUSTIFY,
        spaceAfter=8,
        firstLineIndent=0,
    ))

    styles.add(ParagraphStyle(
        name='Closing',
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=MEDIUM_GRAY,
        alignment=TA_LEFT,
        spaceBefore=4,
        spaceAfter=2,
    ))

    styles.add(ParagraphStyle(
        name='SignatureName',
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=DARK_NAVY,
        alignment=TA_LEFT,
        spaceBefore=4,
    ))

    return styles


def build_cover_letter(filename, role_title, paragraphs, is_template=False):
    """Build a single cover letter PDF."""
    filepath = os.path.join(OUTPUT_DIR, filename)
    doc = SimpleDocTemplate(
        filepath,
        pagesize=letter,
        topMargin=0.6 * inch,
        bottomMargin=0.5 * inch,
        leftMargin=0.85 * inch,
        rightMargin=0.85 * inch,
    )

    styles = get_styles()
    story = []

    # Header: Name
    story.append(Paragraph("Sumesh Chakkaravarthi Purushothaman", styles['ContactName']))

    # Contact info line
    contact_line = (
        "Boston, MA &nbsp;|&nbsp; (857) 351-9440 &nbsp;|&nbsp; "
        "purushothaman.s@northeastern.edu"
    )
    story.append(Paragraph(contact_line, styles['ContactInfo']))

    contact_line2 = (
        '<link href="https://www.linkedin.com/in/sumesh-chakkaravarthi-purushothaman-1236631b7/" color="#2d5f8a">'
        "LinkedIn</link> &nbsp;|&nbsp; "
        '<link href="https://sumesh-chakkaravarthi.github.io/" color="#2d5f8a">'
        "Portfolio</link>"
    )
    story.append(Paragraph(contact_line2, styles['ContactInfo']))

    # Horizontal rule
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1.2, color=LINE_COLOR, spaceAfter=8, spaceBefore=0))

    # Date
    date_text = "[Date]" if is_template else current_date
    story.append(Paragraph(date_text, styles['DateLine']))

    # Hiring info
    if is_template:
        story.append(Paragraph("[Hiring Manager's Name]", styles['DateLine']))
        story.append(Paragraph("[Company Name]", styles['DateLine']))
        story.append(Paragraph("[Company Address]", styles['DateLine']))
    else:
        story.append(Paragraph("Hiring Team", styles['DateLine']))

    story.append(Spacer(1, 6))

    # Salutation
    story.append(Paragraph("Dear Hiring Manager,", styles['Salutation']))

    # Body paragraphs
    for para in paragraphs:
        story.append(Paragraph(para, styles['BodyText_Custom']))

    # Closing
    story.append(Spacer(1, 4))
    story.append(Paragraph("Thank you for your time and consideration.", styles['BodyText_Custom']))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Sincerely,", styles['Closing']))
    story.append(Paragraph("Sumesh Chakkaravarthi Purushothaman", styles['SignatureName']))

    doc.build(story)
    print(f"  ✅ Created: {filepath}")
    return filepath


# ─── Cover Letter Content ─────────────────────────────────────────────────────

DATA_ANALYST_PARAGRAPHS = [
    "I am writing to express my strong interest in the <b>Data Analyst</b> position at your organization. "
    "As a Master's candidate in Analytics (Applied Machine Intelligence) at Northeastern University with a 3.81 GPA, "
    "I bring a rigorous analytical foundation combined with hands-on experience transforming complex datasets into "
    "clear, actionable business insights.",

    "In my role at Power of Patients, I engineered end-to-end analytical workflows integrating multi-source clinical "
    "and NOAA weather datasets using Python, Pandas, and SQL. I designed 7-day rolling temporal features synchronized "
    "by ZIP code, performed statistical analysis using Random Forest and SHAP explainability frameworks, and delivered "
    "reproducible, stakeholder-ready analytics through Jupyter notebooks and dashboards — all while maintaining strict "
    "NDA compliance. At Cisco AICTE, I optimized SQL queries and Excel workflows achieving a 20% efficiency improvement, "
    "automated data processing pipelines reducing manual effort by 30%, and built Tableau dashboards for real-time KPI "
    "monitoring in Agile sprint cycles.",

    "My project portfolio further demonstrates my analytical capabilities: I built a centralized PostgreSQL data warehouse "
    "migrating 5,000+ vendor records and developed interactive Tableau/Power BI dashboards that reduced manual reporting "
    "by 40%; forecasted Medicare prescription demand using ARIMA with MAPE under 8%; analyzed 500,000+ traffic stop records "
    "using statistical testing in R to deliver policy recommendations; and developed an ML fraud detection pipeline "
    "achieving 92% accuracy with comprehensive feature engineering.",

    "I am proficient in <b>Python, SQL, R, Tableau, Power BI, Pandas, NumPy, Scikit-Learn</b>, and experienced with "
    "<b>PostgreSQL, MySQL, AWS, Git, Excel, statistical modeling, A/B testing</b>, and EDA techniques. I thrive in "
    "collaborative, Agile environments and excel at communicating complex findings to both technical and non-technical "
    "stakeholders. I would welcome the opportunity to bring my analytical rigor and passion for data-driven insights "
    "to your team.",
]

ML_ENGINEER_PARAGRAPHS = [
    "I am writing to express my keen interest in the <b>Machine Learning Engineer</b> position at your organization. "
    "Currently pursuing a Master's in Analytics (Applied Machine Intelligence) at Northeastern University with a 3.81 GPA, "
    "I bring hands-on experience building production-grade ML pipelines, deploying scalable AI systems, and translating "
    "complex model outputs into actionable business value.",

    "At Power of Patients, I engineered an end-to-end ML pipeline analyzing clinical and NOAA weather datasets to predict "
    "Sickle Cell Disease pain crises, designing 7-day rolling temporal features synchronized by ZIP code using Python, Pandas, "
    "and NOAA API. I built production-grade Random Forest and Logistic Regression classifiers with Scikit-Learn, implemented "
    "SHAP explainability for model interpretability, and deployed reproducible ML workflows with comprehensive documentation "
    "and version control through Git. At Cisco AICTE, I implemented clustering algorithms for pattern detection and anomaly "
    "identification while optimizing data pipelines for a 30% reduction in manual effort.",

    "My project work includes engineering a scalable RAG-based conversational AI system using LangChain with PostgreSQL "
    "and vector embeddings deployed on Red Hat OpenShift, reducing query latency by 40%; building time series forecasting "
    "models using ARIMA achieving MAPE under 8%; developing an ML fraud detection pipeline with Random Forest achieving "
    "92% accuracy using GridSearchCV and SMOTE; and creating an immersive AR/VR learning platform integrating NLP, "
    "computer vision, and TensorFlow-powered recommendation systems.",

    "My technical stack spans <b>Python, TensorFlow, PyTorch, Scikit-Learn, LangChain, SHAP, Docker, AWS, PostgreSQL, "
    "and OpenShift</b>, with deep expertise in supervised/unsupervised learning, deep learning, NLP, computer vision, "
    "and time series analysis. I am passionate about building intelligent systems that are not only technically robust "
    "but also interpretable and impactful. I would be excited to contribute my ML engineering skills and research-driven "
    "mindset to your team.",
]

DATA_ENGINEER_PARAGRAPHS = [
    "I am writing to express my strong interest in the <b>Data Engineer</b> position at your organization. "
    "As a Master's candidate in Analytics (Applied Machine Intelligence) at Northeastern University with a 3.81 GPA "
    "and a B.Tech in AI and Data Science, I specialize in building reliable data pipelines, reporting systems, and "
    "analytics workflows that power data-driven decision-making.",

    "At Power of Patients, I built and maintained ETL-style data pipelines integrating multi-source clinical and NOAA "
    "weather datasets using Python, Pandas, and SQL, engineering 7-day rolling features by ZIP code and date. I performed "
    "data modeling, transformation, and validation to ensure dataset accuracy and reliability; automated recurring data "
    "preparation tasks reducing manual effort; monitored data quality across ingestion and transformation stages implementing "
    "validation checks; and collaborated with stakeholders to evaluate evolving data requirements and recommend schema and "
    "process improvements. At Cisco AICTE, I developed and automated Python- and SQL-based data processing pipelines, "
    "reducing manual preprocessing by 30% and optimizing query performance by 20%.",

    "In my projects, I engineered a centralized PostgreSQL data warehouse by migrating 5,000+ records from spreadsheets, "
    "enabling standardized reporting and improved data integrity. I built automated SQL-driven reporting pipelines feeding "
    "Tableau and Power BI dashboards, reducing manual reporting effort by 40%. I also processed and validated 500,000+ "
    "traffic stop records using R and SQL-style data transformations, and developed structured data analysis workflows "
    "for transaction-level fraud detection datasets.",

    "I bring strong proficiency in <b>Python, SQL, ETL pipelines, Pandas, PostgreSQL, MySQL, Redshift, AWS (EC2, S3), "
    "Docker, Git</b>, and data validation techniques. I am experienced in data modeling, schema design, pipeline "
    "monitoring, and cross-functional collaboration in Agile environments. I am eager to bring my engineering mindset "
    "and passion for building robust data infrastructure to your team.",
]

SOFTWARE_ENGINEER_PARAGRAPHS = [
    "I am writing to express my strong interest in the <b>Software Engineer</b> position at your organization. "
    "Currently pursuing a Master's in Analytics (Applied Machine Intelligence) at Northeastern University with a 3.81 GPA, "
    "I bring experience in Python and backend systems, with a focus on building scalable data pipelines, batch processing "
    "workflows, and reliable infrastructure for data and ML workloads.",

    "At Power of Patients, I designed and implemented backend data pipelines integrating multi-source clinical and NOAA "
    "weather datasets using Python and SQL, supporting repeatable batch processing and downstream analytics workflows. "
    "I engineered deterministic, time-windowed feature pipelines ensuring reproducibility and consistency across inference "
    "and reporting runs; implemented data modeling, transformation, and validation logic for correctness and reliability; "
    "automated recurring pipeline execution reducing manual intervention; and monitored data quality across ingestion stages "
    "enforcing validation checks. I collaborated with product and analytics stakeholders to translate ambiguous requirements "
    "into actionable pipeline changes. At Cisco AICTE, I developed Python- and SQL-based processing pipelines reducing "
    "manual effort by 30% and optimized query performance by 20%.",

    "My project portfolio includes designing and building a centralized PostgreSQL data store with structured querying and "
    "scalable reporting; implementing automated SQL-driven batch pipelines with defined KPI logic and data refresh semantics; "
    "and engineering containerized microservices on Red Hat OpenShift collaborating with Akamai engineering teams for "
    "enterprise-grade deployment. I also built RESTful API endpoints with comprehensive error handling and implemented "
    "caching strategies and database indexing for improved concurrent user handling.",

    "My technical skills include <b>Python, SQL, Java, Bash, PostgreSQL, MySQL, Redshift, Docker, AWS, OpenShift, Git</b>, "
    "with experience in backend development, API design, batch processing, and CI/CD practices. I thrive in structured, "
    "standards-driven environments and am passionate about writing clean, maintainable code that scales. I would welcome "
    "the opportunity to contribute my engineering skills to your team.",
]

AI_ML_ENGINEER_PARAGRAPHS = [
    "I am writing to express my strong interest in the <b>AI/Machine Learning Engineer</b> position at your organization. "
    "As a Master's candidate in Analytics (Applied Machine Intelligence) at Northeastern University with a 3.81 GPA, "
    "I specialize in building autonomous AI agents, multi-protocol communication bridges, and scalable ML pipelines "
    "that translate complex data into verifiable intelligent systems.",

    "At Power of Patients, I engineered deterministic ML feature pipelines for autonomous health monitoring, processing "
    "clinical and NOAA datasets to predict pain crises with 7-day rolling temporal synchronization. I implemented Random "
    "Forest models using the SHAP explainability framework to provide verifiable reasoning chains for patient outreach, "
    "achieving production-ready interpretability. I deployed reproducible ML workflows implementing end-to-end ML lifecycle "
    "practices including experiment tracking, model versioning, and deployment-ready workflows aligned with MLOps best "
    "practices. At Cisco AICTE, I developed automated data processing pipelines implementing clustering algorithms for "
    "pattern detection and anomaly identification.",

    "My project work includes designing a scalable Autonomous AI Agent using LangChain and RAG with PostgreSQL backend "
    "for 7,000+ educational institutions deployed on Red Hat OpenShift, achieving 40% lower query latency through optimized "
    "vector embeddings; developing RESTful API endpoints and vector-based retrieval pipelines for semantic discovery; building "
    "production-ready time series forecasting models with ARIMA achieving MAPE under 8%; creating end-to-end ML fraud detection "
    "achieving 92% accuracy with explainable AI interpretations; and developing an immersive AR/VR platform integrating NLP, "
    "computer vision, and TensorFlow-powered recommendations.",

    "I bring deep expertise in <b>Python, TensorFlow, PyTorch, Scikit-Learn, LangChain, Docker, AWS, OpenShift, PostgreSQL, "
    "and Vector DBs</b>, spanning supervised/unsupervised learning, deep learning, NLP, computer vision, RAG architectures, "
    "and agentic AI systems. I am passionate about pushing the boundaries of AI — from autonomous agents to decentralized "
    "architectures — and I would be excited to bring my research-driven, systems-level thinking to your team.",
]

GENERIC_TEMPLATE_PARAGRAPHS = [
    "I am writing to express my interest in the <b>[Position Title]</b> at <b>[Company Name]</b>. "
    "As a Master's candidate in Analytics (Applied Machine Intelligence) at Northeastern University with a 3.81 GPA, "
    "I bring a strong foundation in data analytics, machine learning, data engineering, and business intelligence, "
    "combined with hands-on experience translating complex datasets into clear, actionable insights.",

    "Across my professional experiences, I have worked extensively with Python, SQL, Tableau, and Power BI to build "
    "dashboards, engineer data pipelines, perform advanced analysis, and support business decision-making. At Power of "
    "Patients, I designed end-to-end analytical workflows integrating multi-source datasets, engineered temporal features, "
    "and delivered stakeholder-ready insights while maintaining strict data governance compliance. At Cisco AICTE, I "
    "automated data processing pipelines reducing manual effort by 30% and optimized query performance by 20%, delivering "
    "insights through Tableau dashboards in Agile sprint cycles.",

    "My project portfolio includes engineering a centralized PostgreSQL data warehouse with automated Tableau/Power BI "
    "dashboards reducing manual reporting by 40%; building production-ready ML and forecasting models achieving high "
    "accuracy; deploying scalable AI systems on Red Hat OpenShift with containerized microservices; and performing large-scale "
    "statistical analysis delivering actionable policy recommendations. These experiences have strengthened my ability to "
    "design reliable data workflows, build intelligent solutions, and communicate complex findings effectively.",

    "What I bring is not just technical proficiency in <b>[list key relevant skills]</b>, but a strong analytical mindset "
    "and a genuine interest in solving business problems with data. I am comfortable working in structured, Agile environments, "
    "communicating insights to both technical and non-technical stakeholders, and continuously learning to keep pace with "
    "evolving platforms. I would welcome the opportunity to discuss how my skills and experiences can add value to "
    "<b>[Company Name]</b>.",
]


def main():
    print("\n📄 Generating Professional Cover Letters...\n")

    letters = [
        ("Sumesh_Chakkaravarthi_Data_Analyst_Cover_Letter.pdf", "Data Analyst", DATA_ANALYST_PARAGRAPHS, False),
        ("Sumesh_Chakkaravarthi_ML_Engineer_Cover_Letter.pdf", "Machine Learning Engineer", ML_ENGINEER_PARAGRAPHS, False),
        ("Sumesh_Chakkaravarthi_Data_Engineer_Cover_Letter.pdf", "Data Engineer", DATA_ENGINEER_PARAGRAPHS, False),
        ("Sumesh_Chakkaravarthi_Software_Engineer_Cover_Letter.pdf", "Software Engineer", SOFTWARE_ENGINEER_PARAGRAPHS, False),
        ("Sumesh_Chakkaravarthi_AI_ML_Engineer_Cover_Letter.pdf", "AI/ML Engineer", AI_ML_ENGINEER_PARAGRAPHS, False),
        ("Sumesh_Chakkaravarthi_Cover_Letter_Template.pdf", "Generic Template", GENERIC_TEMPLATE_PARAGRAPHS, True),
    ]

    for filename, role, paragraphs, is_template in letters:
        build_cover_letter(filename, role, paragraphs, is_template)

    print(f"\n✨ All cover letters saved to: {OUTPUT_DIR}\n")


if __name__ == "__main__":
    main()
