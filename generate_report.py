"""
Generate the College Event Analytics System – Project Report (.docx)
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CHARTS = "charts"
OUT    = "College_Event_Analytics_Report.docx"

doc = Document()

# ─── Page margins ─────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3.0)
    section.right_margin  = Cm(2.5)

# ─── Helper utilities ─────────────────────────────────────────────────────────

def set_font(run, name="Times New Roman", size=12, bold=False, italic=False, color=None):
    run.font.name  = name
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.italic= italic
    if color:
        run.font.color.rgb = RGBColor(*color)

def heading(text, level=1, center=False, color=(0,0,0)):
    p = doc.add_heading(level=level)
    p.clear()
    run = p.add_run(text)
    set_font(run, size=16 if level==1 else (14 if level==2 else 12),
             bold=True, color=color)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p

def para(text, size=12, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         color=(0,0,0), space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    run = p.add_run(text)
    set_font(run, size=size, bold=bold, italic=italic, color=color)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(18)
    return p

def bullet(text, size=12):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    set_font(run, size=size)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = Pt(18)

def add_image(path, caption, width=Inches(5.8)):
    if os.path.exists(path):
        doc.add_picture(path, width=width)
        last = doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp = doc.add_paragraph(caption)
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cp.runs[0]
        set_font(run, size=10, italic=True, color=(80,80,80))
        cp.paragraph_format.space_after = Pt(10)

def page_break():
    doc.add_page_break()

def hline():
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "4472C4")
    pBdr.append(bottom)
    pPr.append(pBdr)
    p.paragraph_format.space_after = Pt(6)

# ═══════════════════════════════════════════════════════════════════════════════
# TITLE PAGE
# ═══════════════════════════════════════════════════════════════════════════════

doc.add_paragraph()
doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("COLLEGE EVENT ANALYTICS SYSTEM")
set_font(r, size=22, bold=True, color=(31,73,125))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("A Data-Driven Approach to Event Performance Evaluation")
set_font(r, size=14, italic=True, color=(68,114,196))

doc.add_paragraph()
hline()
doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Project Report")
set_font(r, size=16, bold=True)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Submitted in partial fulfilment of the requirements for the\n"
              "Bachelor of Technology / Bachelor of Computer Applications")
set_font(r, size=12)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Department of Computer Science & Engineering")
set_font(r, size=13, bold=True)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Academic Year 2024 – 2025")
set_font(r, size=12)

doc.add_paragraph()
doc.add_paragraph()

hline()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Submitted by")
set_font(r, size=12, italic=True)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("[Your Name]\n[Roll Number]\n[Your College Name]")
set_font(r, size=13, bold=True)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# ACKNOWLEDGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

heading("Acknowledgement", level=1, center=True, color=(31,73,125))
hline()
doc.add_paragraph()

para(
    "First and foremost, I express my sincere gratitude to the Almighty for giving me the "
    "strength, knowledge, and ability to undertake this research project and to persevere "
    "and complete it satisfactorily."
)
para(
    "I would like to extend my heartfelt thanks to my project guide and mentor, "
    "[Mentor's Name], [Designation], Department of Computer Science & Engineering, "
    "for their invaluable guidance, continuous encouragement, and unwavering support "
    "throughout the duration of this project. Their expertise and insightful feedback "
    "have been instrumental in shaping this work."
)
para(
    "I am deeply grateful to the Head of the Department, [HOD Name], and all the faculty "
    "members of the Department of Computer Science & Engineering for their continuous "
    "motivation and for providing the necessary infrastructure and resources that made "
    "this project possible."
)
para(
    "I would also like to sincerely thank the college event organizers and administrative "
    "staff who provided access to the event data – including registration records, "
    "attendance logs, feedback forms, and student profiles – that formed the backbone "
    "of this analytical study."
)
para(
    "Special thanks go to my classmates and peers who offered constructive criticism, "
    "suggestions, and moral support during the development of this project. Their "
    "collaborative spirit greatly enriched the experience."
)
para(
    "Finally, I owe an immense debt of gratitude to my family for their unconditional "
    "love, patience, and encouragement. Their constant motivation gave me the confidence "
    "to pursue this work wholeheartedly."
)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = p.add_run("[Your Name]")
set_font(r, size=12, bold=True)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# ABSTRACT
# ═══════════════════════════════════════════════════════════════════════════════

heading("Abstract", level=1, center=True, color=(31,73,125))
hline()
doc.add_paragraph()

para(
    "The College Event Analytics System is a comprehensive data analytics project designed "
    "to evaluate and improve the planning, execution, and impact of college events. With "
    "the growing scale of student activities in academic institutions — ranging from "
    "technical competitions and cultural fests to workshops, career fairs, and sports meets "
    "— there is an increasing need to move beyond subjective evaluations and adopt "
    "data-driven decision-making approaches."
)
para(
    "This project analyses six interlinked datasets comprising 74,030 registration records "
    "across 400 events involving 6,000 unique students. The datasets include event metadata, "
    "student registrations, attendance logs, feedback ratings, sentiment data, and student "
    "demographic profiles. Using Python and its ecosystem of data science libraries "
    "(pandas, NumPy, Matplotlib, Seaborn, SciPy), the system performs end-to-end analysis "
    "covering data loading, cleaning, preprocessing, exploratory data analysis, statistical "
    "hypothesis testing, and rich data visualisation."
)
para(
    "Key findings reveal an overall attendance rate of 77.4%, with Sports and Technical "
    "categories leading in participation. The most registered event, Hackathon 2023, "
    "attracted the highest footfall. Student satisfaction averaged 3.72 out of 5, with "
    "60.8% positive sentiment. Club membership and early registration were identified as "
    "significant predictors of attendance. Statistical tests (T-test, ANOVA) confirmed "
    "significant differences in attendance across demographic groups. Ten detailed "
    "visualisation charts provide actionable insights into event popularity, student "
    "engagement, satisfaction trends, promotion effectiveness, and timing preferences."
)
para(
    "The system provides event organisers with a structured framework to evaluate past "
    "performance and make informed decisions for future college events, ultimately "
    "enhancing student experience and institutional outcomes."
)

doc.add_paragraph()
para("Keywords: Event Analytics, Data Visualisation, Student Participation, Attendance Analysis, "
     "Feedback Analysis, Exploratory Data Analysis, College Events, Python.", italic=True)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# TABLE OF CONTENTS (Manual)
# ═══════════════════════════════════════════════════════════════════════════════

heading("Table of Contents", level=1, center=True, color=(31,73,125))
hline()
doc.add_paragraph()

toc_items = [
    ("Acknowledgement", "ii"),
    ("Abstract", "iii"),
    ("Chapter 1 – Introduction", "1"),
    ("    1.1  Background and Motivation", "1"),
    ("    1.2  Problem Statement", "2"),
    ("    1.3  Objectives", "2"),
    ("    1.4  Scope of the Project", "3"),
    ("    1.5  Tools and Technologies Used", "3"),
    ("Chapter 2 – Dataset Selection and Data Preparation", "4"),
    ("    2.1  Dataset Overview", "4"),
    ("    2.2  Data Loading", "5"),
    ("    2.3  Data Cleaning", "5"),
    ("    2.4  Feature Engineering", "6"),
    ("    2.5  Missing Value Analysis", "6"),
    ("Chapter 3 – Exploratory Data Analysis and Statistical Analysis", "7"),
    ("    3.1  Descriptive Statistics", "7"),
    ("    3.2  Category-Wise Analysis", "7"),
    ("    3.3  Feedback & Sentiment Analysis", "8"),
    ("    3.4  Statistical Hypothesis Testing", "8"),
    ("Chapter 4 – Data Visualization and Insights", "9"),
    ("    4.1  Overview Dashboard", "9"),
    ("    4.2  Event Popularity", "10"),
    ("    4.3  Participation Trends", "11"),
    ("    4.4  Satisfaction & Feedback", "12"),
    ("    4.5  Event Performance", "13"),
    ("    4.6  Student Profiles", "14"),
    ("    4.7  Correlation Heatmap", "15"),
    ("    4.8  Promotion & Timing", "16"),
    ("    4.9  Top & Bottom Performers", "17"),
    ("    4.10 Registration Lead-Time", "18"),
    ("Chapter 5 – Conclusion and Future Scope", "19"),
    ("    5.1  Conclusion", "19"),
    ("    5.2  Future Scope", "20"),
    ("References", "21"),
]

tbl = doc.add_table(rows=len(toc_items), cols=2)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (item, pg) in enumerate(toc_items):
    row = tbl.rows[i]
    c0, c1 = row.cells[0], row.cells[1]
    c0.width = Inches(5.2)
    c1.width = Inches(0.8)
    r0 = c0.paragraphs[0].add_run(item)
    r1 = c1.paragraphs[0].add_run(pg)
    c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    bold = not item.startswith(" ")
    set_font(r0, size=11, bold=bold)
    set_font(r1, size=11, bold=bold)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# CHAPTER 1 – INTRODUCTION
# ═══════════════════════════════════════════════════════════════════════════════

heading("Chapter 1 – Introduction", level=1, color=(31,73,125))
hline()

heading("1.1  Background and Motivation", level=2, color=(68,114,196))
para(
    "College events are a cornerstone of student life in academic institutions. They serve "
    "as platforms for intellectual exchange, cultural expression, skill development, and "
    "community building. From hackathons and technical symposia to cultural festivals, "
    "career fairs, and sports meets, events provide students with opportunities that go far "
    "beyond the traditional classroom environment."
)
para(
    "Despite the critical role these events play in student development, most college event "
    "organizers rely heavily on informal observations, word-of-mouth feedback, and manual "
    "attendance counts to evaluate event success. This approach is inherently subjective, "
    "error-prone, and incapable of capturing the nuanced patterns that drive student "
    "participation and satisfaction."
)
para(
    "The advent of digital registration systems, automated attendance tracking, and online "
    "feedback platforms has generated a wealth of structured data about college events. "
    "Harnessing this data through modern analytics techniques offers enormous potential "
    "to transform the way events are planned and evaluated. The College Event Analytics "
    "System is conceived precisely to fill this gap — translating raw event data into "
    "clear, actionable intelligence for organizers and administrators."
)

heading("1.2  Problem Statement", level=2, color=(68,114,196))
para(
    "Despite the availability of digital data from college event management systems, "
    "there is currently no systematic, data-driven approach to evaluating event performance "
    "in most institutions. Organizers lack visibility into:\n"
    "  • Which event categories drive the highest student engagement.\n"
    "  • What factors (timing, promotion, fees, certificates) influence attendance.\n"
    "  • How satisfaction levels vary across event types and demographics.\n"
    "  • Which events are underperforming and need redesign.\n"
    "  • What student profiles are most and least engaged.\n\n"
    "Without structured analysis, resources continue to be misallocated, and the quality "
    "of events stagnates rather than improving year over year."
)

heading("1.3  Objectives", level=2, color=(68,114,196))
para("The primary objectives of this project are:")
bullet("To load, clean, and preprocess multi-source college event datasets into a unified analytical framework.")
bullet("To perform comprehensive Exploratory Data Analysis (EDA) to uncover patterns in registrations, attendance, and feedback.")
bullet("To conduct statistical hypothesis testing to validate findings about student engagement.")
bullet("To create 10 detailed data visualisations covering all major dimensions of event analytics.")
bullet("To derive actionable insights and recommendations for improving future college event planning.")
bullet("To present a structured, reproducible analytics pipeline using Python and open-source tools.")

heading("1.4  Scope of the Project", level=2, color=(68,114,196))
para(
    "This project covers the analysis of college event data spanning two academic years "
    "(2023–2024 and 2024–2025), encompassing 400 events, 6,000 students, and over 74,000 "
    "registration records. The analysis covers seven event categories: Technical, Cultural, "
    "Workshop, Career, Sports, Seminar, and Social/NSS. The system is designed to be "
    "generalisable and can be adapted to analyse event data from any college or university "
    "with similar data structures."
)
para(
    "The scope includes data preparation, statistical analysis, and visualisation. "
    "It does not include real-time event management, predictive modelling, or web "
    "deployment — these are identified as future scope items."
)

heading("1.5  Tools and Technologies Used", level=2, color=(68,114,196))
para("The following tools and technologies were employed in this project:")

tools_data = [
    ("Tool / Library", "Version", "Purpose"),
    ("Python",         "3.14",    "Core programming language"),
    ("pandas",         "2.x",     "Data loading, manipulation, and aggregation"),
    ("NumPy",          "2.x",     "Numerical computations and array operations"),
    ("Matplotlib",     "3.x",     "Base charting and figure composition"),
    ("Seaborn",        "0.13.x",  "Statistical and aesthetic visualisations"),
    ("SciPy",          "1.x",     "Statistical hypothesis testing (T-test, ANOVA)"),
    ("python-docx",    "1.x",     "Programmatic Word document generation"),
    ("VS Code / IDE",  "–",       "Development environment"),
    ("Git",            "–",       "Version control"),
]

tbl2 = doc.add_table(rows=len(tools_data), cols=3)
tbl2.style = "Table Grid"
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, row_data in enumerate(tools_data):
    row = tbl2.rows[i]
    for j, val in enumerate(row_data):
        cell = row.cells[j]
        run = cell.paragraphs[0].add_run(val)
        set_font(run, size=11, bold=(i == 0))
        if i == 0:
            cell._tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "4472C4")
            cell._tc.tcPr.append(shd)
            run.font.color.rgb = RGBColor(255, 255, 255)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# CHAPTER 2 – DATASET SELECTION AND DATA PREPARATION
# ═══════════════════════════════════════════════════════════════════════════════

heading("Chapter 2 – Dataset Selection and Data Preparation", level=1, color=(31,73,125))
hline()

heading("2.1  Dataset Overview", level=2, color=(68,114,196))
para(
    "The project utilises six interrelated CSV datasets that together capture the full "
    "lifecycle of college event activities — from event creation and student registration "
    "through attendance and post-event feedback. Together, these datasets form a rich, "
    "multi-dimensional view of student engagement."
)

ds_data = [
    ("Dataset",           "File",                 "Rows",   "Columns", "Description"),
    ("Master Dataset",    "master_dataset.csv",   "74,030", "45",      "Unified view joining all datasets"),
    ("Events",            "events.csv",           "400",    "25",      "Event metadata, category, timing, features"),
    ("Registrations",     "registrations.csv",    "74,030", "8",       "Student registration records per event"),
    ("Attendance",        "attendance.csv",        "57,313", "6",       "Check-in logs with time and duration"),
    ("Feedback",          "feedback.csv",          "34,013", "11",      "Ratings, sentiments, and comments"),
    ("Students",          "students.csv",          "6,000",  "7",       "Student demographics, CGPA, club membership"),
]
tbl3 = doc.add_table(rows=len(ds_data), cols=5)
tbl3.style = "Table Grid"
tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, row_data in enumerate(ds_data):
    row = tbl3.rows[i]
    for j, val in enumerate(row_data):
        cell = row.cells[j]
        run = cell.paragraphs[0].add_run(val)
        set_font(run, size=10, bold=(i == 0))
        if i == 0:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"),   "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"),  "4472C4")
            cell._tc.get_or_add_tcPr().append(shd)
            run.font.color.rgb = RGBColor(255,255,255)

doc.add_paragraph()

heading("2.2  Data Loading", level=2, color=(68,114,196))
para(
    "All six datasets were loaded using the Python pandas library. Each file was read into "
    "a separate DataFrame and then inspected for shape, column names, data types, and "
    "initial null counts. The master_dataset.csv serves as the primary analytical table, "
    "having been pre-joined across events, registrations, attendance, feedback, and "
    "student records. The individual files provide granular detail for targeted analyses."
)
para(
    "Data loading confirmed: 74,030 registration entries across 400 events, with "
    "57,313 attendance records (students who actually showed up) and 34,013 feedback "
    "submissions. The master dataset contains 45 columns covering every dimension of "
    "the event lifecycle."
)

heading("2.3  Data Cleaning", level=2, color=(68,114,196))
para(
    "Data cleaning addressed the following issues identified during the initial inspection:"
)
bullet("Date parsing: event_date and registration_date columns were converted from string "
       "format to pandas datetime objects to enable temporal analysis.")
bullet("Missing guest_speaker_rating: 30.6% of events had no speaker rating (events with "
       "no external guest speaker). These were imputed with the column median (3.8) to "
       "preserve rows for downstream analysis.")
bullet("Text standardisation: Category, event_type, gender, residence, and sentiment "
       "columns were stripped of leading/trailing whitespace and converted to Title Case "
       "for consistency.")
bullet("Duplicate checks: No duplicate registration_id entries were found in the master "
       "dataset, confirming data integrity.")
bullet("Feedback missing values: 54.1% of rating columns (overall_rating, "
       "organization_rating, content_rating, venue_rating, would_recommend, sentiment) "
       "were null — corresponding to the 40,017 students who registered but did not "
       "submit feedback. These were retained as-is, as excluding them would distort "
       "attendance and registration analyses.")

heading("2.4  Feature Engineering", level=2, color=(68,114,196))
para(
    "Several new features were derived from the cleaned data to support richer analysis:"
)
bullet("event_month, event_year, event_quarter: Extracted from event_date to enable "
       "temporal trend analysis.")
bullet("attendance_rate (per event): Calculated as (total_attended / total_registrations) × 100, "
       "merged back onto the events table.")
bullet("avg_overall_rating, avg_organization_rating, avg_content_rating, avg_venue_rating: "
       "Mean feedback scores aggregated at event level.")
bullet("cgpa_band: CGPA binned into five bands (<5, 5-6, 6-7, 7-8, 8+) to analyse the "
       "relationship between academic performance and event participation.")
bullet("composite_score: A blended metric combining attendance rate (50%) and average "
       "overall rating scaled to 100 (50%) to rank event performance.")
bullet("leadtime_band: Registration days_before_event binned into five ranges to study "
       "early vs. late registration behaviour.")
bullet("promo_band: promotion_days binned to study the effect of promotion duration on "
       "attendance.")

heading("2.5  Missing Value Analysis", level=2, color=(68,114,196))
mv_data = [
    ("Column",               "Missing Count", "Missing %", "Treatment"),
    ("guest_speaker_rating", "22,661",        "30.6%",     "Imputed with median (3.8)"),
    ("overall_rating",       "40,017",        "54.1%",     "Retained (no feedback submitted)"),
    ("organization_rating",  "40,017",        "54.1%",     "Retained"),
    ("content_rating",       "40,017",        "54.1%",     "Retained"),
    ("venue_rating",         "40,017",        "54.1%",     "Retained"),
    ("would_recommend",      "40,017",        "54.1%",     "Retained"),
    ("sentiment",            "40,017",        "54.1%",     "Retained"),
    ("comment",              "48,523",        "65.5%",     "Excluded from rating analysis"),
]
tbl4 = doc.add_table(rows=len(mv_data), cols=4)
tbl4.style = "Table Grid"
tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, row_data in enumerate(mv_data):
    row = tbl4.rows[i]
    for j, val in enumerate(row_data):
        cell = row.cells[j]
        run = cell.paragraphs[0].add_run(val)
        set_font(run, size=10, bold=(i == 0))
        if i == 0:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"),   "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"),  "4472C4")
            cell._tc.get_or_add_tcPr().append(shd)
            run.font.color.rgb = RGBColor(255,255,255)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# CHAPTER 3 – EDA AND STATISTICAL ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════

heading("Chapter 3 – Exploratory Data Analysis and Statistical Analysis", level=1, color=(31,73,125))
hline()

heading("3.1  Descriptive Statistics", level=2, color=(68,114,196))
para(
    "The following high-level statistics summarise the dataset at a glance:"
)
stats_data = [
    ("Metric",                        "Value"),
    ("Total unique events",           "400"),
    ("Total unique students",         "5,994"),
    ("Total registration records",    "74,030"),
    ("Total attendance records",      "57,313"),
    ("Overall attendance rate",       "77.42%"),
    ("Total feedback submissions",    "34,013"),
    ("Average overall rating",        "3.72 / 5"),
    ("Would-recommend rate",          "60.8%"),
    ("Positive sentiment share",      "60.8%"),
    ("Best day for attendance",       "Saturday"),
    ("Best time slot",                "Afternoon"),
    ("Most registered event",         "Hackathon 2023"),
]
tbl5 = doc.add_table(rows=len(stats_data), cols=2)
tbl5.style = "Table Grid"
tbl5.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(stats_data):
    row = tbl5.rows[i]
    rk = row.cells[0].paragraphs[0].add_run(k)
    rv = row.cells[1].paragraphs[0].add_run(v)
    set_font(rk, size=11, bold=(i == 0))
    set_font(rv, size=11, bold=(i == 0))
    if i == 0:
        for cell in row.cells:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"),   "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"),  "4472C4")
            cell._tc.get_or_add_tcPr().append(shd)
            for run in cell.paragraphs[0].runs:
                run.font.color.rgb = RGBColor(255,255,255)
    elif i % 2 == 0:
        for cell in row.cells:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"),   "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"),  "DCE6F1")
            cell._tc.get_or_add_tcPr().append(shd)

doc.add_paragraph()

heading("3.2  Category-Wise Analysis", level=2, color=(68,114,196))
para(
    "Registrations are distributed unevenly across the seven event categories, reflecting "
    "the varying appeal and frequency of different event types:"
)
cat_data = [
    ("Category",    "Registrations", "Attendance Rate"),
    ("Technical",   "25,154",        "78.95%"),
    ("Cultural",    "13,275",        "78.31%"),
    ("Workshop",    "10,298",        "77.08%"),
    ("Career",      "10,064",        "71.52%"),
    ("Sports",      "6,524",         "78.99%"),
    ("Seminar",     "5,853",         "77.77%"),
    ("Social/NSS",  "2,862",         "77.50%"),
]
tbl6 = doc.add_table(rows=len(cat_data), cols=3)
tbl6.style = "Table Grid"
tbl6.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, row_data in enumerate(cat_data):
    row = tbl6.rows[i]
    for j, val in enumerate(row_data):
        cell = row.cells[j]
        run = cell.paragraphs[0].add_run(val)
        set_font(run, size=11, bold=(i == 0))
        if i == 0:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"),   "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"),  "4472C4")
            cell._tc.get_or_add_tcPr().append(shd)
            run.font.color.rgb = RGBColor(255,255,255)

doc.add_paragraph()
para(
    "Technical events dominate registrations (33.9% share), driven by the high frequency "
    "and appeal of hackathons, tech talks, and coding competitions among the student body. "
    "However, Sports events record the highest attendance rate (79%), suggesting that "
    "students who register for sports are more committed to attending. Career events show "
    "the lowest conversion from registration to attendance (71.5%), indicating potential "
    "mismatches between student expectations and actual career event delivery."
)

heading("3.3  Feedback and Sentiment Analysis", level=2, color=(68,114,196))
para(
    "Of the 74,030 registrations, 34,013 students (45.9%) submitted post-event feedback. "
    "The average ratings across the four dimensions were:"
)
rating_data = [
    ("Dimension",            "Mean Rating", "Std Dev"),
    ("Overall Rating",       "3.72",        "0.87"),
    ("Organisation Rating",  "3.70",        "0.96"),
    ("Content Rating",       "3.70",        "0.96"),
    ("Venue Rating",         "3.91",        "0.81"),
]
tbl7 = doc.add_table(rows=len(rating_data), cols=3)
tbl7.style = "Table Grid"
tbl7.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, row_data in enumerate(rating_data):
    row = tbl7.rows[i]
    for j, val in enumerate(row_data):
        cell = row.cells[j]
        run = cell.paragraphs[0].add_run(val)
        set_font(run, size=11, bold=(i == 0))
        if i == 0:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"),   "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"),  "4472C4")
            cell._tc.get_or_add_tcPr().append(shd)
            run.font.color.rgb = RGBColor(255,255,255)

doc.add_paragraph()
para(
    "Venue quality received the highest average rating (3.91), suggesting students are "
    "generally satisfied with the physical infrastructure of events. Content and organisation "
    "scores (3.70 each) indicate room for improvement in event quality and management. "
    "The sentiment distribution showed 60.8% positive, 31.8% neutral, and 7.4% negative "
    "feedback — broadly encouraging but with scope to convert neutral responses to positive."
)

heading("3.4  Statistical Hypothesis Testing", level=2, color=(68,114,196))
para(
    "Two formal statistical tests were conducted to validate analytical observations:"
)
para("Independent Samples T-Test – Club Members vs. Non-Members:", bold=True, size=12)
para(
    "Null Hypothesis (H₀): There is no significant difference in attendance rates between "
    "club members and non-members.\n"
    "Result: t = 38.79, p < 0.0001\n"
    "Conclusion: The null hypothesis is strongly rejected. Club members attend college events "
    "at a significantly higher rate than non-members. This finding suggests that fostering "
    "club culture and encouraging club memberships can be a direct lever for improving "
    "event attendance."
)
para("One-Way ANOVA – Attendance Across Departments:", bold=True, size=12)
para(
    "Null Hypothesis (H₀): Attendance rates are equal across all academic departments.\n"
    "Result: F = 2.47, p = 0.0083\n"
    "Conclusion: The null hypothesis is rejected at a 1% significance level. There are "
    "statistically significant differences in event attendance across departments. "
    "This implies that targeted engagement strategies may need to be department-specific, "
    "addressing the unique barriers and motivations of each student group."
)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# CHAPTER 4 – DATA VISUALISATION AND INSIGHTS
# ═══════════════════════════════════════════════════════════════════════════════

heading("Chapter 4 – Data Visualization and Insights", level=1, color=(31,73,125))
hline()

para(
    "This chapter presents ten data visualisation charts produced by the College Event "
    "Analytics System. Each chart is accompanied by a detailed interpretation of the "
    "patterns and insights it reveals. All charts were generated using Matplotlib and "
    "Seaborn libraries in Python and are saved at 150 DPI resolution."
)

# 4.1
heading("4.1  Overview Dashboard", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "01_overview_dashboard.png"),
          "Figure 4.1 – Overview Dashboard: Registrations by Category, Attendance Rate, "
          "Sentiment Distribution, Monthly Trends, and Registration Mode")
para(
    "The overview dashboard provides a bird's-eye view of the entire event ecosystem. "
    "The top-left panel confirms Technical events as the dominant category with the highest "
    "registration count, followed by Cultural and Workshop. The attendance rate panel "
    "shows Sports with the highest conversion rate (~79%), while Career events lag behind "
    "at ~71.5%."
)
para(
    "The sentiment pie chart reveals a largely positive student experience — 60.8% positive, "
    "31.8% neutral, and only 7.4% negative — indicating that events are broadly well-received "
    "but there is consistent room for improvement, especially in converting neutral respondents."
)
para(
    "The monthly registration trend line shows clear seasonal spikes, with peaks typically "
    "occurring at the start of semesters (January and August) when event activity is highest. "
    "Troughs in April–May and October–November likely correspond to examination periods "
    "where student availability drops. The registration mode pie confirms that Club Referrals "
    "and Online portals are the dominant channels, with walk-in registrations forming a "
    "smaller share."
)

# 4.2
heading("4.2  Event Popularity Analysis", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "02_event_popularity.png"),
          "Figure 4.2 – Top 15 Events by Registrations and Capacity Fill Rate")
para(
    "The event popularity chart highlights the top 15 most-registered events and their "
    "respective capacity fill rates. Hackathon 2023 emerges as the most popular event "
    "by registration count, consistent with the high student interest in technical "
    "competitions that offer tangible rewards such as certificates, prizes, and networking "
    "opportunities."
)
para(
    "The capacity fill rate panel reveals that several events exceeded or came very close "
    "to their maximum capacity (100% line marked in red), indicating strong demand that "
    "organisers should anticipate in future planning by increasing venue capacity or "
    "running multiple batches. Conversely, events with low fill rates despite adequate "
    "promotion may require restructuring of content or scheduling to improve appeal."
)

# 4.3
heading("4.3  Participation Trends", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "03_participation_trends.png"),
          "Figure 4.3 – Attendance Rate by Department, Year of Study, Club Membership, and Gender")
para(
    "The participation trends analysis reveals several important demographic patterns "
    "in student event engagement:"
)
para(
    "By Department: Attendance rates vary meaningfully across departments, with the "
    "ANOVA test confirming statistical significance (F=2.47, p=0.0083). Some departments "
    "show consistently higher participation, possibly due to club culture, peer influence, "
    "or proximity of events to departmental interests."
)
para(
    "By Year of Study: First and second-year students tend to have higher event participation "
    "rates, which is consistent with the pattern of newer students exploring campus life "
    "more actively. Senior students (3rd and 4th year) may be more selective in their "
    "participation due to academic pressures and placement activities."
)
para(
    "By Club Membership: Club members attend events at a dramatically higher rate than "
    "non-members (T-test: t=38.79, p<0.0001). This is the single strongest predictor "
    "of attendance in the dataset, highlighting the role of student clubs in driving "
    "event culture and participation."
)
para(
    "By Gender: Both male and female students show similar overall attendance rates, "
    "but some category-specific differences emerge. Cultural events attract relatively "
    "more female participation, while Technical and Sports events show stronger male "
    "attendance. These insights can guide targeted marketing for underrepresented groups."
)

# 4.4
heading("4.4  Satisfaction and Feedback Analysis", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "04_satisfaction_feedback.png"),
          "Figure 4.4 – Average Ratings, Category Satisfaction, Would-Recommend Rate, and Sentiment Trend")
para(
    "The satisfaction analysis provides a comprehensive view of student experience across "
    "multiple dimensions:"
)
para(
    "Average ratings across the four dimensions (Overall, Organisation, Content, Venue) "
    "cluster around 3.7–3.9 out of 5. Venue receives the highest satisfaction score "
    "(3.91), reflecting generally adequate physical infrastructure. Content and Organisation "
    "scores (3.70 each) are the weakest links, suggesting that improving event content "
    "quality, speaker selection, and organisational logistics would have the greatest "
    "impact on student satisfaction."
)
para(
    "The category-wise rating chart shows that Sports and Cultural events tend to receive "
    "the highest satisfaction scores, while Seminars and Career events score relatively "
    "lower — possibly due to high expectations versus actual utility perceived by students."
)
para(
    "The 'Would Recommend' rate by category is strongly correlated with overall ratings, "
    "confirming internal consistency. Events with >65% recommendation rates are prime "
    "candidates for scale-up and promotion in future semesters."
)
para(
    "The quarterly sentiment trend area chart shows the evolution of positive, neutral, "
    "and negative sentiment over time. Positive sentiment is consistently dominant, "
    "but shows mild dips in exam-adjacent quarters, suggesting that event quality or "
    "student mood is influenced by the academic calendar."
)

# 4.5
heading("4.5  Event Performance Deep-Dive", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "05_event_performance.png"),
          "Figure 4.5 – Free vs Paid Events, Time Slot Analysis, and Duration vs Rating Scatter")
para(
    "The event performance deep-dive examines structural event characteristics and their "
    "relationship with attendance and satisfaction:"
)
para(
    "Free vs. Paid Events: Free events show a higher average attendance rate than paid "
    "events, which is expected given the financial barrier of a registration fee. "
    "However, the margin is smaller than anticipated, suggesting that students are willing "
    "to pay for high-value events with strong content, prizes, or certificates."
)
para(
    "Time Slot Analysis: Afternoon sessions show the highest average attendance rate, "
    "followed by Morning sessions. Evening events have the lowest participation, likely "
    "because students face transportation and hostel curfew constraints. This finding "
    "provides a clear scheduling recommendation: schedule high-priority events in the "
    "afternoon slot to maximise turnout."
)
para(
    "Duration vs. Rating Scatter: The scatter plot reveals a mild positive correlation "
    "between event duration and average rating — longer events tend to receive marginally "
    "higher satisfaction scores, possibly because longer formats allow for more content "
    "depth, networking, or entertainment. Events with both high attendance (warm colours) "
    "and high ratings cluster in the 3-6 hour duration range, suggesting this is the "
    "sweet spot for event length."
)

# 4.6
heading("4.6  Student Profile Analysis", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "06_student_profiles.png"),
          "Figure 4.6 – Attendance by CGPA Band, Residence Type, Department Distribution, and CGPA Histogram")
para(
    "The student profile analysis investigates how demographic and academic characteristics "
    "correlate with event participation:"
)
para(
    "CGPA and Attendance: Students in the higher CGPA bands (7-8 and 8+) show slightly "
    "higher attendance rates than lower CGPA students. This counter-intuitive finding "
    "(one might expect high-achievers to prioritise academics over events) suggests that "
    "high-performing students are well-rounded and manage both academics and extracurricular "
    "activities effectively. However, the differences are relatively modest."
)
para(
    "Residence and Attendance: Hostel residents show a higher event attendance rate than "
    "Day Scholars. This is attributed to the greater physical proximity of hostel students "
    "to the campus, lower commute burden, and stronger peer influence in residential "
    "environments. Event organisers can leverage hostel student networks as word-of-mouth "
    "ambassadors."
)
para(
    "Department Distribution: The student population spans multiple departments, with "
    "some departments (e.g., CSE, ECE) having significantly larger student counts. "
    "This size advantage translates into higher absolute registration numbers from "
    "those departments, though per-capita participation rates tell a more equitable story."
)
para(
    "CGPA Distribution: The CGPA histogram approximates a near-normal distribution "
    "centred around 6.5–7.5, with most students in the average-to-good academic "
    "performance band. Very few students fall below 5 or above 9, confirming that "
    "the dataset is representative of a typical college population."
)

# 4.7
heading("4.7  Feature Correlation Heatmap", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "07_correlation_heatmap.png"),
          "Figure 4.7 – Correlation Heatmap of Numeric Event and Student Features")
para(
    "The correlation heatmap provides a simultaneous view of relationships between all "
    "numeric variables in the dataset. Key observations include:"
)
bullet("Promotion channels and promotion days show a moderate positive correlation, "
       "indicating that events with longer promotion periods tend to also use more channels.")
bullet("Capacity and duration show a mild positive correlation — larger-capacity events "
       "tend to be longer events.")
bullet("Club membership shows a meaningful positive correlation with attended, "
       "confirming the statistical test results.")
bullet("Near_exams shows a negative correlation with attendance, confirming that events "
       "scheduled close to examination periods attract fewer students.")
bullet("The heatmap confirms the absence of severe multicollinearity between independent "
       "variables, validating the statistical integrity of analyses conducted.")

# 4.8
heading("4.8  Promotion and Timing Analysis", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "08_promotion_timing.png"),
          "Figure 4.8 – Attendance by Promotion Duration, Day of Week, and Exam Proximity")
para(
    "Effective promotion and scheduling are two of the most controllable factors influencing "
    "event attendance. The promotion and timing analysis reveals:"
)
para(
    "Promotion Duration: Events promoted for 14–21 days before the event show the highest "
    "attendance rates. Very short promotion windows (<7 days) result in lower awareness "
    "and hence lower attendance. Surprisingly, very long promotion windows (30+ days) "
    "also show slightly reduced effectiveness, possibly due to reminder fatigue or "
    "information overload. The optimal promotion window appears to be 2–3 weeks."
)
para(
    "Day of Week: Saturday events attract the highest attendance rates, followed by "
    "Sunday events. Weekday events, especially Monday and Friday, record lower participation "
    "— consistent with student preferences for weekend activities and the challenge of "
    "attending events on days with heavy academic schedules. Organisers should prioritise "
    "weekend scheduling for flagship events."
)
para(
    "Exam Proximity: Events scheduled within the exam period show a clear and significant "
    "drop in attendance compared to non-exam period events. This underscores the importance "
    "of aligning the college event calendar with the academic schedule to avoid conflicts "
    "with examination windows."
)

# 4.9
heading("4.9  Top and Bottom Performing Events", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "09_top_bottom_events.png"),
          "Figure 4.9 – Top 10 and Bottom 10 Events by Composite Performance Score")
para(
    "A composite performance score was calculated for each event, combining attendance "
    "rate (50% weight) and average overall rating scaled to 100 (50% weight). This "
    "balanced metric identifies events that are both well-attended and well-received."
)
para(
    "Top Performers: The top 10 events share common characteristics — they are typically "
    "weekend events, have certificate/prize incentives, were promoted for 2–3 weeks, "
    "and belong to the Technical or Cultural categories. Hackathon 2023 and similar "
    "competitive events dominate the top of the ranking."
)
para(
    "Bottom Performers: The bottom 10 events tend to be weekday afternoon events "
    "with fewer incentives (no prizes or certificates), shorter promotion windows, "
    "or were scheduled close to examination periods. Several Career category events "
    "appear in the lower tier due to lower-than-expected attendance and moderate "
    "satisfaction scores. These events require significant redesign in terms of "
    "content, timing, and promotion strategy."
)

# 4.10
heading("4.10  Registration Lead-Time Analysis", level=2, color=(68,114,196))
add_image(os.path.join(CHARTS, "10_registration_leadtime.png"),
          "Figure 4.10 – Registration Lead-Time Distribution and Attendance by Lead-Time Band")
para(
    "The registration lead-time analysis examines the relationship between how far in "
    "advance students register and whether they ultimately attend the event:"
)
para(
    "Lead-Time Distribution: The histogram shows that most students register within "
    "0–10 days of the event, with a peak at 3–7 days. Very early registrations (30+ days "
    "ahead) are relatively rare, suggesting that students make event participation decisions "
    "close to the event date. This has implications for organisers — early registration "
    "campaigns and incentives (early-bird discounts, guaranteed spots) can help secure "
    "attendance commitments earlier."
)
para(
    "Attendance Rate by Lead-Time: Students who register 8–14 days before the event "
    "show the highest attendance rate, suggesting that this window represents a 'committed' "
    "registration — long enough to plan ahead but close enough to remain motivated. "
    "Last-minute registrations (0–3 days) show lower attendance rates, indicating "
    "impulsive registrations that often don't convert to attendance. Organisers should "
    "consider confirmation reminders for last-minute registrants."
)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# CHAPTER 5 – CONCLUSION AND FUTURE SCOPE
# ═══════════════════════════════════════════════════════════════════════════════

heading("Chapter 5 – Conclusion and Future Scope", level=1, color=(31,73,125))
hline()

heading("5.1  Conclusion", level=2, color=(68,114,196))
para(
    "The College Event Analytics System successfully demonstrates the power of data-driven "
    "analysis in evaluating and improving college event management. By integrating six "
    "interrelated datasets and applying a rigorous analytical pipeline — spanning data "
    "loading, cleaning, preprocessing, EDA, statistical hypothesis testing, and rich "
    "visualisation — the system delivers a comprehensive and nuanced understanding of "
    "college event dynamics."
)
para(
    "The key takeaways from this study are:"
)
bullet("The overall attendance rate of 77.4% across 400 events and 74,030 registrations "
       "reflects a generally healthy event participation culture in the institution.")
bullet("Technical events are the most popular by registration count, but Sports events "
       "lead in attendance conversion rate, suggesting stronger commitment among sports "
       "enthusiasts.")
bullet("Student satisfaction is broadly positive (average rating 3.72/5, 60.8% positive "
       "sentiment), with venue quality the highest-rated dimension and content/organisation "
       "identified as the primary areas for improvement.")
bullet("Club membership is the strongest single predictor of event attendance "
       "(T-test: t=38.79, p<0.0001), making club development a high-leverage strategy "
       "for improving event participation.")
bullet("Afternoon sessions on weekends, events promoted 14–21 days in advance, "
       "and events not scheduled near exams consistently outperform others in attendance.")
bullet("Early-to-moderate registration lead times (8–14 days) correlate with higher "
       "attendance rates, while last-minute registrations often don't convert.")
bullet("Statistical tests confirm that attendance differences across departments and "
       "between club/non-club members are not due to chance — they reflect real "
       "structural differences that organisers should address.")
para(
    "In conclusion, this project establishes a replicable, scalable framework for "
    "institutional event analytics that can empower college administrators and student "
    "organisations to make better, evidence-based decisions for future events."
)

heading("5.2  Future Scope", level=2, color=(68,114,196))
para(
    "While the current system delivers significant value, several enhancements can "
    "extend its capabilities further:"
)
bullet("Predictive Analytics: Machine learning models (Random Forest, Logistic Regression, "
       "XGBoost) can be trained to predict individual student attendance probability "
       "before an event, enabling targeted outreach and personalised reminders.")
bullet("Natural Language Processing: Sentiment analysis on the free-text comment column "
       "using NLP techniques (BERT, VADER) would surface specific issues and praise "
       "that quantitative ratings cannot capture.")
bullet("Real-Time Dashboard: A web-based interactive dashboard (using Dash, Streamlit, "
       "or Power BI) would allow organisers to monitor live registration and attendance "
       "data during events and access post-event analytics instantly.")
bullet("Recommendation System: An event recommendation engine could suggest relevant "
       "upcoming events to students based on their past participation history and "
       "academic profile, improving personalised engagement.")
bullet("Multi-Institution Benchmarking: Extending the system to aggregate data across "
       "multiple colleges or universities would enable benchmarking of event performance "
       "relative to peer institutions.")
bullet("Integration with College ERP: Direct integration with the institution's "
       "Enterprise Resource Planning (ERP) or Student Information System (SIS) would "
       "automate data collection and eliminate manual CSV exports.")
bullet("A/B Testing Framework: Implementing structured A/B tests for event parameters "
       "(different time slots, pricing strategies, promotion channels) would enable "
       "rigorous causal inference rather than correlational analysis.")

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# REFERENCES
# ═══════════════════════════════════════════════════════════════════════════════

heading("References", level=1, color=(31,73,125))
hline()
doc.add_paragraph()

refs = [
    "[1]  McKinney, W. (2010). Data Structures for Statistical Computing in Python. "
         "Proceedings of the 9th Python in Science Conference, 51-56.",
    "[2]  Hunter, J. D. (2007). Matplotlib: A 2D Graphics Environment. Computing in "
         "Science & Engineering, 9(3), 90-95.",
    "[3]  Waskom, M. (2021). Seaborn: Statistical Data Visualization. Journal of Open "
         "Source Software, 6(60), 3021.",
    "[4]  Virtanen, P., et al. (2020). SciPy 1.0: Fundamental Algorithms for Scientific "
         "Computing in Python. Nature Methods, 17, 261-272.",
    "[5]  Harris, C. R., et al. (2020). Array Programming with NumPy. Nature, 585, "
         "357-362.",
    "[6]  Van Rossum, G., & Drake, F. L. (2009). Python 3 Reference Manual. Scotts Valley, "
         "CA: CreateSpace.",
    "[7]  Tukey, J. W. (1977). Exploratory Data Analysis. Addison-Wesley.",
    "[8]  Field, A. (2013). Discovering Statistics Using IBM SPSS Statistics (4th ed.). "
         "SAGE Publications.",
    "[9]  Wickham, H. (2010). A Layered Grammar of Graphics. Journal of Computational "
         "and Graphical Statistics, 19(1), 3-28.",
    "[10] python-docx Documentation. https://python-docx.readthedocs.io/ (Accessed: Oct 2026).",
    "[11] pandas Documentation. https://pandas.pydata.org/docs/ (Accessed: Oct 2026).",
    "[12] Matplotlib Documentation. https://matplotlib.org/stable/contents.html (Accessed: Oct 2026).",
]

for ref in refs:
    p = doc.add_paragraph()
    r = p.add_run(ref)
    set_font(r, size=11)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = Pt(18)

# ─── Save ─────────────────────────────────────────────────────────────────────
doc.save(OUT)
print(f"\nReport saved: {OUT}")
