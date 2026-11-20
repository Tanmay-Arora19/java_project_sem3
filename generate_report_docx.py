import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

doc = Document()

# Set standard 1-inch margins
for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Header and Footer
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run("Chitkara University | Department of CSE (AI/ML)")
    hrun.font.name = "Times New Roman"
    hrun.font.size = Pt(8.5)
    hrun.font.italic = True
    hrun.font.color.rgb = RGBColor(100, 116, 139)
    
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    frun = fp.add_run("Student Performance Analytics Tool — Semester III Project Report")
    frun.font.name = "Times New Roman"
    frun.font.size = Pt(8.5)
    frun.font.color.rgb = RGBColor(100, 116, 139)

# Colors
COLOR_PRIMARY = RGBColor(11, 37, 69)     # Deep Navy
COLOR_SECONDARY = RGBColor(19, 64, 116)  # Royal Blue
COLOR_ACCENT = RGBColor(0, 119, 182)     # Academic Blue
COLOR_BODY = RGBColor(45, 55, 72)        # Charcoal Slate
COLOR_MUTED = RGBColor(100, 116, 139)

# Helper function to style paragraphs
def add_para(text, font_size=11, bold=False, italic=False, color=COLOR_BODY, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6, line_spacing=1.15):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return p

def add_heading_1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY
    return p

def add_heading_2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = COLOR_SECONDARY
    return p

def add_heading_3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY
    return p

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def style_table(table, col_widths, headers, rows):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Header Row
    hdr_cells = table.rows[0].cells
    for i, h_text in enumerate(headers):
        hdr_cells[i].text = h_text
        set_cell_background(hdr_cells[i], "0B2545")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=140, right=140)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            
    # Data Rows
    for r_idx, row_data in enumerate(rows):
        row_cells = table.add_row().cells
        bg_color = "F8FAFC" if (r_idx % 2 == 1) else "FFFFFF"
        for c_idx, cell_value in enumerate(row_data):
            row_cells[c_idx].text = str(cell_value)
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=90, bottom=90, left=140, right=140)
            p = row_cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(9.5)
                run.font.color.rgb = COLOR_BODY
                
    # Set explicit column widths
    for row in table.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)

def add_captioned_image(img_path, caption, width_inches=6.0):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run = p_img.add_run()
        run.add_picture(img_path, width=Inches(width_inches))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(12)
        r_cap = p_cap.add_run(caption)
        r_cap.font.name = "Times New Roman"
        r_cap.font.size = Pt(9.5)
        r_cap.font.italic = True
        r_cap.font.color.rgb = COLOR_MUTED

def add_code_block(code_text):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F1F5F9")
    set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
    
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(code_text)
    run.font.name = "Consolas"
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(15, 23, 42)
    
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(6)

# =============================================================
# COVER PAGE / TITLE PAGE
# =============================================================
add_para("CHITKARA UNIVERSITY", font_size=18, bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
add_para("DEPARTMENT OF CSE (AI/ML)", font_size=14, bold=True, color=COLOR_SECONDARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)

add_captioned_image("images/fig4_1_architecture.png", "System Architecture & High-Level Design Overview", width_inches=5.8)

add_para("STUDENT PERFORMANCE ANALYTICS TOOL", font_size=22, bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
add_para("A Console-Based Core Java Academic Evaluation & Performance Tracking System", font_size=12, italic=True, color=COLOR_ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

add_para(
    "A modular, robust academic evaluation system built with Core Java, Object-Oriented Programming (OOP), "
    "Java Collections Framework, Java Streams API, custom exception handling, multithreaded notifications, and "
    "file persistence with automated byte-level disaster recovery backups.",
    font_size=10.5, italic=False, color=COLOR_BODY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20
)

add_para("Submitted by (Group Members):", font_size=11, bold=True, color=COLOR_PRIMARY, space_after=6)

team_table = doc.add_table(rows=1, cols=3)
team_headers = ["S. No.", "Name of Student", "Roll Number"]
team_rows = [
    ["1", "Tanmay Arora", "2510993577"],
    ["2", "Ansh Arora", "2510992560"],
    ["3", "Jaskaran Singh Bhinder", "2510993488"],
    ["4", "Vishal", "2510992990"]
]
style_table(team_table, [0.8, 3.2, 2.5], team_headers, team_rows)

p_guide = doc.add_paragraph()
p_guide.paragraph_format.space_before = Pt(24)
p_guide.paragraph_format.space_after = Pt(4)
p_guide.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_guide = p_guide.add_run("Semester III | Chitkara University, Punjab\nUnder the guidance of: Rohit Sharma\nAcademic Year 2026")
r_guide.font.name = "Times New Roman"
r_guide.font.size = Pt(11)
r_guide.font.bold = True
r_guide.font.color.rgb = COLOR_PRIMARY

doc.add_page_break()

# =============================================================
# BONAFIDE CERTIFICATE
# =============================================================
add_para("Bonafide Certificate", font_size=18, bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

add_para(
    'This is to certify that the project report titled "Student Performance Analytics Tool" is a bonafide record '
    'of group project work carried out jointly by the team members named below, students of B.Tech Computer Science & '
    'Engineering (AI/ML), under my supervision and guidance, in partial fulfilment of the requirements for the '
    'Java Programming course project during Semester III.',
    font_size=11, space_after=10
)

add_para(
    "Each member listed below contributed to the design, development and documentation of this project, and shares "
    "joint responsibility for the work presented in this report.",
    font_size=11, space_after=14
)

cert_table = doc.add_table(rows=1, cols=3)
style_table(cert_table, [0.8, 3.2, 2.5], team_headers, team_rows)

p_sign = doc.add_paragraph()
p_sign.paragraph_format.space_before = Pt(45)
p_sign.paragraph_format.space_after = Pt(25)
r1 = p_sign.add_run("_______________________________\t\t\t_______________________________\n")
r1.font.bold = True
r2 = p_sign.add_run("Team Representative Signature\t\t\tFaculty Guide Signature\n")
r2.font.bold = True
r3 = p_sign.add_run("On behalf of all group members\t\t\tRohit Sharma, CSE (AI/ML)")
r3.font.color.rgb = COLOR_MUTED

p_hod = doc.add_paragraph()
p_hod.paragraph_format.space_before = Pt(25)
p_hod.paragraph_format.space_after = Pt(10)
rh1 = p_hod.add_run("_______________________________\n")
rh1.font.bold = True
rh2 = p_hod.add_run("Head of Department\nDepartment of CSE (AI/ML), Chitkara University")
rh2.font.bold = True

doc.add_page_break()

# =============================================================
# ACKNOWLEDGEMENT
# =============================================================
add_para("Acknowledgement", font_size=18, bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

add_para(
    "We, the members of this project group, would like to express our sincere gratitude to our faculty guide, Rohit Sharma, "
    "for the valuable guidance, encouragement and constructive feedback provided throughout the duration of this project. "
    "Their support helped us understand not just how to write Java code, but how to architect a modular, maintainable console application "
    "using Core Java, Object-Oriented design principles, Java Collections, and modern Stream API pipelines.",
    font_size=11, space_after=10
)

add_para(
    "We are also thankful to the Department of Computer Science & Engineering (AI/ML), Chitkara University, for providing the "
    "resources, computer laboratory facilities and structured coursework that made it possible to translate academic concepts "
    "into a complete, working software project.",
    font_size=11, space_after=10
)

add_para(
    "This project was a genuinely collaborative effort. Each member contributed across requirements analysis, class design, "
    "exception handling, stream analytics, unit testing, and documentation. Regular team discussions helped us maintain "
    "clean coding standards, resolve edge cases, and divide the tool's modules without losing structural integrity. We are grateful "
    "for each other's commitment, patience, and teamwork throughout this semester.",
    font_size=11, space_after=10
)

add_para(
    "Finally, we acknowledge the global Java open-source community and the documentation provided by Oracle, whose standard "
    "specifications for Java SE 8+ made it possible for our team to independently research and implement robust solutions.",
    font_size=11, space_after=16
)

ack_table = doc.add_table(rows=1, cols=3)
ack_headers = ["S. No.", "Team Member", "Key Contribution Area"]
ack_rows = [
    ["1", "Tanmay Arora", "Main Program Architecture, Switch-Case Controller, File I/O Persistence & Integration"],
    ["2", "Ansh Arora", "Performance Analysis, Java Stream API Analytics, Sorting & Student Ranking Engine"],
    ["3", "Jaskaran Singh Bhinder", "Dynamic Subject Setup, Student Registration, Input Validation & Exception Handling"],
    ["4", "Vishal", "Core Student Class Model, Calculations, Grading Logic & Reportable Interface Contract"]
]
style_table(ack_table, [0.8, 2.2, 3.5], ack_headers, ack_rows)

p_group_close = doc.add_paragraph()
p_group_close.paragraph_format.space_before = Pt(20)
r_close = p_group_close.add_run("On behalf of the project team\nB.Tech CSE (AI/ML), Semester III, Chitkara University")
r_close.font.bold = True
r_close.font.color.rgb = COLOR_PRIMARY

doc.add_page_break()

# =============================================================
# ABSTRACT
# =============================================================
add_para("Abstract", font_size=18, bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

add_para(
    "Managing student academic performance records manually using traditional paper mark registers or unstandardized spreadsheets "
    "is an inefficient, error-prone, and labor-intensive task for teachers and academic evaluators. Manual computation of total marks, "
    "percentages, academic letter grades, merit rankings, and class-wide performance metrics frequently introduces human calculation errors "
    "and consumes valuable faculty time. The Student Performance Analytics Tool is a comprehensive terminal-based application developed "
    "using Core Java (JDK 8+) to eliminate these operational bottlenecks through end-to-end automation.",
    font_size=11, space_after=10
)

add_para(
    "The application incorporates dynamic subject configuration, enabling academic coordinators to specify custom subject catalogs "
    "for any class batch. Student details, including roll numbers, student names, and individual subject scores, are captured and strictly "
    "validated against business rules: duplicate roll numbers are prohibited, and marks are validated within the legal boundary of 0 to 100 "
    "using a custom InvalidMarksException. The system encapsulates student records inside an object-oriented Student data model implementing "
    "the Reportable interface, computing total marks, exact percentage scores, standard letter grades (A+ to F), and identifying each student's "
    "strongest and weakest subjects.",
    font_size=11, space_after=10
)

add_para(
    "For advanced academic evaluation, the tool features a student ranking engine powered by Java's Comparator interface, performing automated "
    "descending-order sorting with alphabetical tie-breaking. It provides subject-wise statistical reporting (class average, highest scorer, and "
    "lowest scorer per subject) and high-level class performance analytics implemented via the Java Stream API, delivering the batch mean percentage, "
    "pass/fail counts, class topper details, and frequency-based grade distribution maps. Data persistence is achieved via character streams "
    "(students.txt), accompanied by an automated byte-stream cloning mechanism (students_backup.txt) for disaster recovery.",
    font_size=11, space_after=12
)

add_para(
    "Keywords: Student Performance Analytics, Core Java, Object-Oriented Programming (OOP), Collections Framework, Stream API, "
    "Custom Exception, Comparator, File I/O, Multithreading, Disaster Recovery.",
    font_size=10.5, bold=True, color=COLOR_PRIMARY, space_after=16
)

doc.add_page_break()

# =============================================================
# TABLE OF CONTENTS
# =============================================================
add_para("Table of Contents", font_size=18, bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

toc_table = doc.add_table(rows=1, cols=3)
toc_headers = ["Chapter #", "Chapter Title", "Page"]
toc_rows = [
    ["1", "Introduction (Background, Problem Statement, Objectives, Scope)", "6"],
    ["2", "Literature Review / Existing Systems (Registers, Spreadsheets, Architectural Gap)", "8"],
    ["3", "System Requirements (Hardware, Software, Technology Stack)", "9"],
    ["4", "System Design (Architecture, OOP Principles, Data Structures)", "11"],
    ["5", "Module-Wise Description (Terminal Outputs & Step-by-Step Walkthrough)", "13"],
    ["6", "Implementation Details (Class Hierarchy, Algorithms & Code Highlights)", "20"],
    ["7", "Testing & Verification (Test Cases Table)", "22"],
    ["8", "Limitations and Future Scope", "23"],
    ["9", "Conclusion", "24"],
    ["10", "References", "25"]
]
style_table(toc_table, [1.0, 4.7, 0.8], toc_headers, toc_rows)

doc.add_page_break()

# =============================================================
# CHAPTER 1: INTRODUCTION
# =============================================================
add_heading_1("CHAPTER 1: Introduction")

add_heading_2("1.1 Background and Motivation")
add_para(
    "In academic institutions, evaluating and tracking student performance across multiple subjects is a continuous requirement. "
    "At the conclusion of each examination cycle, teachers must gather raw scores, verify roll numbers, compute total marks and percentages, "
    "assign institutional letter grades, generate merit ranks, and prepare statistical class summaries for departmental review. "
    "Historically, this process was conducted manually using paper registers or general-purpose spreadsheet software.",
    font_size=11
)
add_para(
    "However, manual calculation of scores for batches of 50 to 100 students is inherently tedious and susceptible to human calculation error. "
    "A minor transposition error during percentage computation or manual sorting can distort class ranks and compromise grading fairness. "
    "The Student Performance Analytics Tool was conceived to demonstrate how Core Java and Object-Oriented Programming (OOP) can be applied "
    "to construct an automated, reliable, and platform-independent evaluation tool that enforces strict data validation and generates instant "
    "analytical insights.",
    font_size=11
)

add_heading_2("1.2 Problem Statement")
add_para("Design and develop a terminal-based Java application that:", font_size=11)
add_para("• Allows teachers to configure dynamic subject lists for an academic class batch.", font_size=11)
add_para("• Captures student roll numbers and names while strictly preventing duplicate registrations.", font_size=11)
add_para("• Enforces score integrity by rejecting values outside the 0–100 boundary via custom exception handling.", font_size=11)
add_para("• Automatically calculates total marks, percentage scores, and official letter grades (A+ through F).", font_size=11)
add_para("• Identifies individual subject strengths (highest score) and weaknesses (lowest score) per student.", font_size=11)
add_para("• Generates fair merit rankings using Java Comparator in descending order with alphabetical tie-breaking.", font_size=11)
add_para("• Produces subject-level analytics (class average, highest score with topper name, lowest score) and class summaries via Stream API.", font_size=11)
add_para("• Preserves all academic records in permanent text storage with an automated byte-stream disaster recovery backup.", font_size=11)

add_heading_2("1.3 Objectives")
add_para("1. Object-Oriented Modeling: Encapsulate student state and calculation behaviors within an autonomous Student entity.", font_size=11)
add_para("2. Input Validation & Fault Tolerance: Prevent application crashes by trapping invalid inputs with custom exceptions.", font_size=11)
add_para("3. Automated Grading: Eliminate human error by calculating exact percentages and evaluating conditional grade branches.", font_size=11)
add_para("4. Algorithmic Ranking: Utilize Java's Comparator interface to sort student records objectively in memory.", font_size=11)
add_para("5. Functional Stream Processing: Aggregate class performance metrics using the modern Java Stream API and Collectors.", font_size=11)
add_para("6. Dual-Stream Persistence: Combine character streams for human-readable data with byte streams for automated backups.", font_size=11)

add_heading_2("1.4 Scope of the Project")
add_para(
    "The boundaries of the Student Performance Analytics Tool within the scope of the 3rd Semester Java Programming curriculum "
    "are outlined in the comparative table below:",
    font_size=11
)

scope_table = doc.add_table(rows=1, cols=2)
scope_headers = ["In Scope (Implemented & Fully Working)", "Out of Scope (Future Roadmap)"]
scope_rows = [
    ["Modular Core Java console application with switch-case menu", "Graphical User Interface (GUI) via JavaFX or Swing"],
    ["Dynamic subject configuration and duplicate roll prevention", "Relational database integration (MySQL / PostgreSQL via JDBC)"],
    ["Custom InvalidMarksException enforcing 0–100 score range", "Web-based cloud portal accessible via HTTP/REST APIs"],
    ["Automatic percentage, letter grading, and highest/lowest subject detection", "Role-based authentication with password hashing and JWT sessions"],
    ["Automated student ranking using Comparator with name tie-breakers", "Automated PDF report card export with official university watermarks"],
    ["Subject-wise statistical analytics and Stream API class summary", "SMS / Email notification gateways for sending results to parents"],
    ["Character stream persistence (students.txt) & byte backup (students_backup.txt)", "Distributed cloud database replication across multi-tenant servers"]
]
style_table(scope_table, [3.2, 3.3], scope_headers, scope_rows)

doc.add_page_break()

# =============================================================
# CHAPTER 2: LITERATURE REVIEW / EXISTING SYSTEMS
# =============================================================
add_heading_1("CHAPTER 2: Literature Review / Existing Systems")

add_heading_2("2.1 Manual Paper Registers")
add_para(
    "For decades, academic mark entry was performed using physical mark registers. Under this paradigm, teachers wrote student names, "
    "recorded exam scores in ruled columns, manually summed rows using hand calculators, and looked up grade boundaries on printed tables. "
    "This method presents severe limitations: human fatigue leads to calculation errors, ranking 60+ students requires sorting slips of paper manually, "
    "and paper registers are highly vulnerable to physical deterioration, misplacement, and ink damage.",
    font_size=11
)

add_heading_2("2.2 Spreadsheets vs. Specialized Tools")
add_para(
    "Many academic institutions migrated from paper to spreadsheet software like Microsoft Excel or Google Sheets. While spreadsheets provide "
    "grid layouts and formula cells (=SUM, =AVERAGE), they lack domain-specific data integrity safeguards: users can accidentally overwrite cell formulas, "
    "enter invalid marks (e.g., negative numbers or values above 100) without immediate error prevention, or break sorting ranges. Furthermore, "
    "complex statistics like subject-wise toppers and grade frequency distributions require complicated multi-cell array formulas that general users "
    "struggle to configure reliably.",
    font_size=11
)

add_heading_2("2.3 The Architectural Gap and Our Java Approach")
add_para(
    "To resolve the tension between fragile spreadsheets and tedious paper registers, the Student Performance Analytics Tool provides a specialized, "
    "rule-enforced Java application. By encapsulating academic rules inside compiled Java bytecode, the application guarantees that data cannot be entered "
    "without passing strict validation routines. The program combines fast in-memory Java Collections (ArrayList and HashMap), declarative sorting via "
    "Comparator, functional Stream API aggregations, and persistent file I/O into a unified, zero-dependency console tool that can run on any system "
    "equipped with a standard Java Runtime Environment (JRE).",
    font_size=11
)

# =============================================================
# CHAPTER 3: SYSTEM REQUIREMENTS
# =============================================================
add_heading_1("CHAPTER 3: System Requirements")

add_heading_2("3.1 Hardware Requirements")
hw_table = doc.add_table(rows=1, cols=2)
hw_headers = ["Hardware Component", "Minimum Specification"]
hw_rows = [
    ["Processor", "Intel Core i3 / AMD Ryzen 3 / Apple Silicon (Dual-core 2.0 GHz or higher)"],
    ["RAM", "2 GB minimum (4 GB recommended for concurrent IDE development)"],
    ["Storage", "100 MB free hard drive space for JDK, source code, and data files"],
    ["Display", "Standard monitor (1024x768 resolution minimum) with command terminal support"],
    ["Input Devices", "Standard physical keyboard for console data entry and menu navigation"]
]
style_table(hw_table, [2.5, 4.0], hw_headers, hw_rows)

add_heading_2("3.2 Software Requirements")
sw_table = doc.add_table(rows=1, cols=2)
sw_headers = ["Software Category", "Environment / Tool Used"]
sw_rows = [
    ["Operating System", "Microsoft Windows 10/11, Linux (Ubuntu/Debian), or macOS"],
    ["Java Development Kit (JDK)", "Oracle OpenJDK 8 or higher (Tested on Java 17 and Java 21)"],
    ["Command Shell / Terminal", "Windows PowerShell, Command Prompt (cmd), or Unix Bash"],
    ["Integrated Development Environment", "Visual Studio Code, IntelliJ IDEA Community, or Eclipse"],
    ["Version Control", "Git version control system and GitHub Cloud Repository"]
]
style_table(sw_table, [2.5, 4.0], sw_headers, sw_rows)

add_heading_2("3.3 Technology Stack & Java Components")
add_para(
    "• Core Java (Java SE): The primary programming language utilized, ensuring strict type safety, modular structure, and platform independence.\n"
    "• Object-Oriented Programming (OOP): Implementation of Encapsulation (private fields with public accessors), Abstraction, Interfaces (Reportable), and Polymorphism.\n"
    "• Java Collections Framework: ArrayList<Student> for dynamic record lists, HashMap<String, Integer> for subject-mark pairings, and TreeMap for sorted grade distributions.\n"
    "• Custom Exception Handling: InvalidMarksException inherits from java.lang.Exception to trap out-of-range mark inputs.\n"
    "• Java Comparator: Comparator.comparingDouble(Student::getPercentage).reversed().thenComparing(Student::getName) for automated merit ranking.\n"
    "• Java Streams & Lambdas: Functional pipelines (.stream(), .mapToDouble(), .filter(), .average(), .max(), Collectors.groupingBy()) for batch analytics.\n"
    "• Multithreading (java.lang.Thread): Background thread execution for asynchronous console notifications when student records are saved.\n"
    "• File I/O Streams: Character streams (FileWriter, PrintWriter, FileReader, BufferedReader) for delimited text persistence and Byte streams (FileInputStream, FileOutputStream) for backup cloning.",
    font_size=10.5
)

doc.add_page_break()

# =============================================================
# CHAPTER 4: SYSTEM DESIGN
# =============================================================
add_heading_1("CHAPTER 4: System Design")

add_heading_2("4.1 System Architecture and Data Flow")
add_para(
    "The Student Performance Analytics Tool is architected around a modular controller-model-view console pattern. The Main Controller "
    "displays an interactive switch-case menu to the teacher. User requests route to specialized input handlers, where raw keyboard inputs "
    "are validated before constructing Student model entities. Business logic (percentage calculation, grade assignment, subject comparison) "
    "is encapsulated directly within the Student class. Output formatting is governed by the Reportable interface, while persistence and backup "
    "services handle disk storage.",
    font_size=11
)

add_captioned_image("images/fig4_1_architecture.png", "Fig 4.1 - System Architecture & OOP Component Data Flow Diagram", width_inches=6.0)

add_heading_2("4.2 Object-Oriented Design Principles")
add_para(
    "1. Encapsulation: The Student class marks its state variables (rollNumber, name, marks) as private. External components cannot modify "
    "these fields directly; access is granted only via explicit getter methods (getRollNumber(), getName(), getMarks(), getPercentage(), getGrade()).",
    font_size=11
)
add_para(
    "2. Interface Abstraction: The nested Reportable interface declares a clean contract void printReport(). By implementing Reportable, "
    "the Student class guarantees that standard report card formatting is decoupled from the main console menu controller.",
    font_size=11
)
add_para(
    "3. Robust Custom Exceptions: The InvalidMarksException class ensures that score validation is handled through standard Java exception-handling "
    "paradigms (try-catch-throw), providing clear error feedback to users without terminating execution.",
    font_size=11
)

add_heading_2("4.3 In-Memory Data Structures")
add_para(
    "• subjects (ArrayList<String>): Maintains an ordered collection of subject names entered during initial setup.\n"
    "• students (ArrayList<Student>): Houses the entire active student body in memory, providing dynamic resizing as new students enroll.\n"
    "• marks (HashMap<String, Integer>): Encapsulated inside each Student instance, mapping subject name keys to numerical mark values (0–100).\n"
    "• ranked (ArrayList<Student>): Cloned temporary collection created during ranking routines to prevent mutating the primary registration order.",
    font_size=11
)

doc.add_page_break()

# =============================================================
# CHAPTER 5: MODULE-WISE DESCRIPTION
# =============================================================
add_heading_1("CHAPTER 5: Module-Wise Description")

add_heading_2("5.1 Main Menu & Application Startup")
add_para(
    "Upon launching the application, the system automatically checks for the existence of students.txt. If present, loadData() parses "
    "the saved records and reconstructs the in-memory collections. The console then renders the primary 9-option interactive menu.",
    font_size=11
)
add_captioned_image("images/fig5_1_main_menu.png", "Fig 5.1 - Terminal Main Menu and Startup Record Loading", width_inches=5.8)

add_heading_2("5.2 Dynamic Subject Configuration")
add_para(
    "Option 1 enables academic coordinators to define custom subjects for the class. The program prompts for the subject count and queries "
    "each subject's name. Duplicate subject names and blank strings are strictly rejected. Once student records have been registered, subject "
    "modifications are locked to prevent schema corruption.",
    font_size=11
)
add_captioned_image("images/fig5_2_setup_subjects.png", "Fig 5.2 - Dynamic Subject Setup Workflow", width_inches=5.8)

add_heading_2("5.3 Student Registration & Score Validation")
add_para(
    "Option 2 captures new student registrations. The user enters a roll number, which is checked against all existing records using findStudent() "
    "to block duplicates. The student's name is read and verified to ensure it is not blank. The user is then prompted for marks in each configured "
    "subject. If a value outside 0–100 is entered, the program throws and catches InvalidMarksException, displaying an error message and reprompting.",
    font_size=11
)
add_captioned_image("images/fig5_3_add_student.png", "Fig 5.3 - Student Record Input and InvalidMarksException Validation", width_inches=5.8)

add_heading_2("5.4 Individual Student Performance Reporting")
add_para(
    "Immediately following registration, or upon searching by roll number (Option 4), the application invokes the printReport() method of the "
    "Reportable interface. The report card displays roll number, student name, subject-by-subject scores out of 100, aggregate marks, computed percentage, "
    "assigned letter grade, and automatically detects the highest and lowest scoring subjects.",
    font_size=11
)
add_captioned_image("images/fig5_4_student_report.png", "Fig 5.4 - Formatted Student Performance Report Card", width_inches=5.8)

add_heading_2("5.5 View All Records & Search by Roll Number")
add_para(
    "Option 3 invokes the generic utility method displayList(List<T> list), iterating through the students ArrayList and printing each student's "
    "concise summary string. Option 4 allows teachers to query an individual student's comprehensive performance report by providing their numerical roll number.",
    font_size=11
)
add_captioned_image("images/fig5_5_view_search.png", "Fig 5.5 - View All Records and Search by Roll Number", width_inches=5.8)

add_heading_2("5.6 Automated Student Ranking via Comparator")
add_para(
    "Option 5 clones the students list into a temporary ranked list and sorts it using a chained Comparator. The comparator sorts students by percentage "
    "in descending order (.reversed()), using student names as an alphabetical tie-breaker (.thenComparing()). The sorted leaderboard is printed with rank numbers.",
    font_size=11
)
add_captioned_image("images/fig5_6_ranking.png", "Fig 5.6 - Merit Ranking Leaderboard via Java Comparator", width_inches=5.8)

add_heading_2("5.7 Subject-Wise Statistical Analytics")
add_para(
    "Option 6 iterates through each registered subject across all student records. For each subject, the algorithm computes the total score, evaluates the "
    "mean class average, tracks the maximum mark with the top scorer's name, and finds the minimum mark with the lowest scoring student's name.",
    font_size=11
)
add_captioned_image("images/fig5_7_subject_statistics.png", "Fig 5.7 - Subject-Wise Performance Analytics", width_inches=5.8)

add_heading_2("5.8 Class Performance Summary via Java Streams")
add_para(
    "Option 7 employs the Java Stream API to compute high-level batch analytics. It calculates the overall class average percentage using .mapToDouble() "
    "and .average(), identifies the class topper using .max(), tallies passed and failed counts via .filter(), and constructs a sorted TreeMap grade distribution "
    "using Collectors.groupingBy(Student::getGrade, TreeMap::new, Collectors.counting()).",
    font_size=11
)
add_captioned_image("images/fig5_8_class_summary.png", "Fig 5.8 - Class Performance Summary via Stream API", width_inches=5.8)

add_heading_2("5.9 File Persistence & Automated Byte Backup")
add_para(
    "Option 8 writes all subject headers and student marks to students.txt using character streams (PrintWriter and FileWriter). Immediately following "
    "the text write, createBackup() utilizes byte streams (FileInputStream and FileOutputStream) to perform a raw byte-by-byte duplicate into students_backup.txt, "
    "ensuring full disaster recovery.",
    font_size=11
)
add_captioned_image("images/fig5_9_file_backup.png", "Fig 5.9 - File Persistence and Byte-Stream Backup Operations", width_inches=5.8)

doc.add_page_break()

# =============================================================
# CHAPTER 6: IMPLEMENTATION DETAILS
# =============================================================
add_heading_1("CHAPTER 6: Implementation Details")

add_heading_2("6.1 Class Hierarchy & Component Structure")
add_para(
    "• StudentPerformanceAnalytics: The main driver class containing console I/O routines, collection references, menu controllers, and file management.\n"
    "• Reportable (Interface): Nested interface declaring void printReport(), establishing contract for formatted report output.\n"
    "• Student (implements Reportable): Inner entity encapsulating rollNumber, name, and marks HashMap, containing all academic evaluation formulas.\n"
    "• InvalidMarksException (extends Exception): Checked exception thrown when an entered mark falls outside the legitimate range of 0 to 100.",
    font_size=11
)

add_heading_2("6.2 Key Mathematical Formulas & Algorithms")
add_para(
    "1. Percentage Computation:\n"
    "   Percentage = (Total Marks Obtained / (Number of Subjects × 100)) × 100",
    font_size=11
)
add_para(
    "2. Academic Grading Criteria:\n"
    "   • 90% – 100%: Grade A+ (Outstanding)\n"
    "   • 80% – 89.99%: Grade A (Excellent)\n"
    "   • 70% – 79.99%: Grade B (Very Good)\n"
    "   • 60% – 69.99%: Grade C (Good)\n"
    "   • 50% – 59.99%: Grade D (Pass)\n"
    "   • Below 50%: Grade F (Fail / Reappear)",
    font_size=11
)

add_heading_2("6.3 Core Java Source Code Highlights")
add_para("Snippet 1: Custom Exception & Student Model Constructor", font_size=10, bold=True)
add_code_block(
"""static class InvalidMarksException extends Exception {
    InvalidMarksException(String message) {
        super(message);
    }
}

static class Student implements Reportable {
    private int rollNumber;
    private String name;
    private HashMap<String, Integer> marks;

    Student(int rollNumber, String name, HashMap<String, Integer> marks) {
        this.rollNumber = rollNumber;
        this.name = name;
        this.marks = marks;
    }
    // Encapsulated getters & calculation methods
}"""
)

add_para("Snippet 2: Chained Comparator Ranking & Stream Analytics", font_size=10, bold=True)
add_code_block(
"""// Sorting via Java Comparator
ArrayList<Student> ranked = new ArrayList<>(students);
Comparator<Student> comparator = Comparator.comparingDouble(Student::getPercentage)
        .reversed().thenComparing(Student::getName);
Collections.sort(ranked, comparator);

// Class summary aggregation via Stream API
double average = students.stream().mapToDouble(Student::getPercentage).average().orElse(0.0);
Student topper = students.stream().max(Comparator.comparingDouble(Student::getPercentage)).get();
long passed = students.stream().filter(s -> !s.getGrade().equals("F")).count();
Map<String, Long> gradeCounts = students.stream()
        .collect(Collectors.groupingBy(Student::getGrade, TreeMap::new, Collectors.counting()));"""
)

add_para("Snippet 3: File Persistence & Byte-Stream Backup", font_size=10, bold=True)
add_code_block(
"""// Character stream writing (PrintWriter)
try (PrintWriter writer = new PrintWriter(new FileWriter("students.txt"))) {
    writer.print("SUBJECTS");
    for (String subject : subjects) writer.print("|" + clean(subject));
    writer.println();
    for (Student s : students) {
        writer.print("STUDENT|" + s.getRollNumber() + "|" + clean(s.getName()));
        for (String sub : subjects) writer.print("|" + s.getMarks(sub));
        writer.println();
    }
}
// Byte-stream backup (FileInputStream & FileOutputStream)
try (FileInputStream src = new FileInputStream("students.txt");
     FileOutputStream dest = new FileOutputStream("students_backup.txt")) {
    int byteData;
    while ((byteData = src.read()) != -1) dest.write(byteData);
}"""
)

doc.add_page_break()

# =============================================================
# CHAPTER 7: TESTING & VERIFICATION
# =============================================================
add_heading_1("CHAPTER 7: Testing")
add_para(
    "Comprehensive verification testing was conducted across data input validation, calculation accuracy, sorting consistency, "
    "stream aggregation, and file persistence. All test cases passed successfully as recorded in the verification matrix below:",
    font_size=11
)

test_table = doc.add_table(rows=1, cols=4)
test_headers = ["Test Case", "Input / Action Taken", "Expected Behavior", "Status"]
test_rows = [
    ["TC-01: Subject Setup", "Enter 3 subjects: Java, Math, DSA", "Configures subjects in memory and saves header line to students.txt", "Pass"],
    ["TC-02: Duplicate Roll Check", "Enter roll number 101 when 101 exists", "Rejects input with 'This roll number already exists.' warning", "Pass"],
    ["TC-03: Score Range (<0)", "Enter mark -15 for Java", "Throws InvalidMarksException: 'Marks must be between 0 and 100.'", "Pass"],
    ["TC-04: Score Range (>100)", "Enter mark 105 for Java", "Throws InvalidMarksException, traps error and safely reprompts user", "Pass"],
    ["TC-05: Percentage Calculation", "Scores: 88, 92, 85 (Total 265/300)", "Calculates exact percentage: 88.33% and assigns Grade A", "Pass"],
    ["TC-06: Subject Extremes", "Scores: 88 (Java), 92 (Math), 85 (DSA)", "Correctly identifies Highest: Math (92) and Lowest: DSA (85)", "Pass"],
    ["TC-07: Comparator Ranking", "4 students with percentages 88.33, 86.67, 84, 80", "Sorts in descending order; top percentage receives Rank 1", "Pass"],
    ["TC-08: Tie-Breaker Ranking", "Two students with identical 85.00% score", "Breaks tie by sorting student names in alphabetical order", "Pass"],
    ["TC-09: Stream Summary", "Class batch of 4 students", "Computes average, pass count (4), fail count (0), topper, and TreeMap", "Pass"],
    ["TC-10: File Reload", "Restart tool after exiting with Option 9", "loadData() reads students.txt and fully reconstructs all objects", "Pass"],
    ["TC-11: Byte-Stream Backup", "Trigger Option 8", "Creates students_backup.txt with exact identical byte count as students.txt", "Pass"]
]
style_table(test_table, [1.5, 1.8, 2.5, 0.7], test_headers, test_rows)

# =============================================================
# CHAPTER 8: LIMITATIONS AND FUTURE SCOPE
# =============================================================
add_heading_1("CHAPTER 8: Limitations and Future Scope")

add_heading_2("8.1 Current Limitations")
add_para(
    "1. Console Interface: The system currently operates via a command-line interface (CLI). While fast and robust, graphical user interfaces (GUI) are generally more visually engaging for non-technical users.\n"
    "2. Local Text File Persistence: Data is stored in delimited flat files on local disk rather than an enterprise relational database (RDBMS). Flat files are susceptible to concurrency conflicts if accessed simultaneously by multiple network processes.\n"
    "3. Single-User Session: The application is designed for a single evaluator executing on a local machine without multi-user authentication or role-based access control.",
    font_size=11
)

add_heading_2("8.2 Planned Future Enhancements")
future_table = doc.add_table(rows=1, cols=2)
future_headers = ["Enhancement Domain", "Proposed Architectural Upgrade"]
future_rows = [
    ["Graphical User Interface (GUI)", "Build a modern desktop GUI using JavaFX or Swing with interactive data tables and dark-mode styling."],
    ["Relational Database Integration", "Connect to MySQL or PostgreSQL using JDBC (Java Database Connectivity) with PreparedStatements."],
    ["Role-Based Authentication", "Implement secure login portals distinguishing between Faculty Evaluators, Department Heads, and Students."],
    ["Automated PDF Report Card Generation", "Integrate Apache PDFBox or iText to export printable, formatted student report cards with university seals."],
    ["Visual Performance Analytics", "Embed graphical bar charts and pie charts representing class grade distributions and performance trends."],
    ["Cloud Web Portal Scaling", "Migrate business logic to a Spring Boot REST API backend supporting distributed mobile and web clients."]
]
style_table(future_table, [2.5, 4.0], future_headers, future_rows)

doc.add_page_break()

# =============================================================
# CHAPTER 9: CONCLUSION
# =============================================================
add_heading_1("CHAPTER 9: Conclusion")
add_para(
    "The Student Performance Analytics Tool project successfully designed, implemented, and validated an automated, modular, "
    "and robust academic record management application built with Core Java. By replacing outdated paper registers and error-prone spreadsheets "
    "with compiled Java business logic, the application completely eliminates human calculation mistakes in percentage scoring, grade assignment, "
    "and student ranking.",
    font_size=11
)
add_para(
    "Through the development of this project, comprehensive practical mastery was achieved across essential Java software engineering disciplines:\n"
    "• Object-Oriented Programming: Encapsulating state in the Student class, adhering to interface contracts with Reportable, and designing clean abstraction boundaries.\n"
    "• Collections & Data Structures: Utilizing ArrayList and HashMap to maintain in-memory subject catalogs and mark mappings efficiently.\n"
    "• Robust Error Handling: Designing and enforcing custom checked exceptions (InvalidMarksException) to trap out-of-range inputs before state mutation.\n"
    "• Algorithmic Sorting: Employing chained Java Comparators to produce fair merit leaderboards with automatic alphabetical tie-breaking.\n"
    "• Modern Functional Streams: Harnessing Java 8+ Stream API pipelines and Collectors to aggregate batch averages, pass/fail counts, and grade frequencies in single expressive statements.\n"
    "• Dual-Stream Persistence: Integrating formatted character streams for text data with raw byte streams for disaster recovery backups.",
    font_size=11
)
add_para(
    "The resulting application provides academic coordinators with a reliable, zero-dependency tool that simplifies semester-end evaluations "
    "and provides a solid architectural foundation ready to be scaled into a full-fledged enterprise university management system.",
    font_size=11
)

# =============================================================
# CHAPTER 10: REFERENCES
# =============================================================
add_heading_1("CHAPTER 10: References")
refs = [
    "1. Oracle Corporation - Official Java Documentation (Java SE 8 & 17 Standard Specifications). https://docs.oracle.com/en/java/",
    "2. Bruce Eckel - 'Thinking in Java', 4th Edition, Prentice Hall PTR, Object-Oriented Principles and Collections Architecture.",
    "3. Herbert Schildt - 'Java: The Complete Reference', 12th Edition, McGraw-Hill Education, Core Java, Streams, and Multithreading.",
    "4. Joshua Bloch - 'Effective Java', 3rd Edition, Addison-Wesley Professional, Best Practices for Class Design, Exceptions, and Streams.",
    "5. Baeldung Java Guides - Comprehensive Tutorials on Java Comparator, Collections Framework, and Stream API Aggregations. https://www.baeldung.com/",
    "6. GeeksforGeeks - Java File Handling, Character Streams vs. Byte Streams, and Custom Exception Design. https://www.geeksforgeeks.org/java/",
    "7. Chitkara University Academic Regulations - Department of Computer Science & Engineering (AI/ML) Project Guidelines and Evaluation Standards."
]
for ref in refs:
    add_para(ref, font_size=10, space_after=6)

# Save Document
output_docx_path = "Student_Performance_Analytics_Report.docx"
doc.save(output_docx_path)
print(f"Report document successfully generated: {output_docx_path}")
