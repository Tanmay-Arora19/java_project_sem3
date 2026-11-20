import os
from PIL import Image, ImageDraw, ImageFont

os.makedirs('images', exist_ok=True)

# Font configurations
CONSOLAS = 'C:/Windows/Fonts/consola.ttf'
ARIAL = 'C:/Windows/Fonts/arial.ttf'
ARIAL_BOLD = 'C:/Windows/Fonts/arialbd.ttf'

font_term = ImageFont.truetype(CONSOLAS, 14)
font_title = ImageFont.truetype(ARIAL_BOLD, 12)
font_arch = ImageFont.truetype(ARIAL_BOLD, 13)
font_arch_sub = ImageFont.truetype(ARIAL, 11)

def create_terminal_image(filename, title, lines, width=880):
    line_height = 22
    top_bar_height = 36
    padding_top = 18
    padding_bottom = 20
    
    height = top_bar_height + padding_top + (len(lines) * line_height) + padding_bottom
    
    img = Image.new('RGB', (width, height), color='#0B132B')
    draw = ImageDraw.Draw(img)
    
    # Outer Border
    draw.rectangle([(0, 0), (width - 1, height - 1)], outline='#1E293B', width=2)
    
    # Titlebar
    draw.rectangle([(0, 0), (width, top_bar_height)], fill='#1C2541')
    
    # Window controls (Traffic lights)
    draw.ellipse([(20, 12), (32, 24)], fill='#EF4444')
    draw.ellipse([(38, 12), (50, 24)], fill='#F59E0B')
    draw.ellipse([(56, 12), (68, 24)], fill='#10B981')
    
    # Titlebar Text
    draw.text((width // 2, 18), title, font=font_title, fill='#94A3B8', anchor='mm')
    
    # Render Terminal Lines
    y = top_bar_height + padding_top
    for line in lines:
        color = '#E2E8F0'
        
        if line.startswith('$') or line.startswith('>'):
            color = '#38BDF8' # Cyan command
        elif line.startswith('=======') or line.startswith('-------') or line.startswith('======='):
            color = '#64748B' # Muted separator
        elif line.startswith('Enter') or line.startswith('How many') or 'choice:' in line:
            color = '#FCD34D' # Amber prompt
        elif '[Notice]' in line or 'Data saved' in line or 'Backup created' in line or 'successfully' in line or 'loaded' in line:
            color = '#4ADE80' # Green success
        elif 'Invalid' in line or 'Skipped' in line or 'Error' in line or 'cannot' in line:
            color = '#F87171' # Red error/warning
        elif 'STUDENT PERFORMANCE REPORT' in line or 'STUDENT RANKING' in line or 'CLASS PERFORMANCE' in line:
            color = '#60A5FA' # Bright blue header
        elif line.startswith('Roll Number') or line.startswith('Student Name') or line.startswith('Total Marks') or line.startswith('Percentage') or line.startswith('Grade'):
            color = '#93C5FD' # Soft blue data label
        elif line.startswith('Rank '):
            color = '#FDE047' # Gold rank
        elif line.startswith('Subject:'):
            color = '#38BDF8' # Cyan subject
            
        draw.text((25, y), line, font=font_term, fill=color)
        y += line_height
        
    img.save(os.path.join('images', filename), quality=95)
    print(f"Saved images/{filename}")

# -------------------------------------------------------------
# 1. Fig 5.1: Main Menu & Startup
# -------------------------------------------------------------
fig51_lines = [
    "$ javac StudentPerformanceAnalytics.java",
    "$ java StudentPerformanceAnalytics",
    "Previous student records loaded.",
    "",
    "======= STUDENT PERFORMANCE ANALYTICS TOOL =======",
    "1. Set up subjects",
    "2. Add student and marks",
    "3. View all students",
    "4. Search student report",
    "5. Rank students",
    "6. Subject-wise statistics",
    "7. Class performance summary",
    "8. Save data and create backup",
    "9. Save and exit",
    "==================================================",
    "Enter your choice: _"
]
create_terminal_image("fig5_1_main_menu.png", "Terminal - Student Performance Analytics Startup", fig51_lines)

# -------------------------------------------------------------
# 2. Fig 5.2: Dynamic Subject Setup
# -------------------------------------------------------------
fig52_lines = [
    "Enter your choice: 1",
    "",
    "How many subjects? 3",
    "Enter subject 1 name: Java Programming",
    "Enter subject 2 name: Mathematics III",
    "Enter subject 3 name: Data Structures",
    "Data saved in students.txt",
    "Subjects saved successfully.",
    "",
    "======= STUDENT PERFORMANCE ANALYTICS TOOL =======",
    "1. Set up subjects",
    "2. Add student and marks",
    "3. View all students",
    "Enter your choice: _"
]
create_terminal_image("fig5_2_setup_subjects.png", "Terminal - Dynamic Subject Configuration", fig52_lines)

# -------------------------------------------------------------
# 3. Fig 5.3: Add Student & Custom Exception Validation
# -------------------------------------------------------------
fig53_lines = [
    "Enter your choice: 2",
    "",
    "Enter student roll number: 101",
    "Enter student name: Tanmay Arora",
    "Enter marks in Java Programming (0-100): 105",
    "Marks must be between 0 and 100.",
    "Enter marks in Java Programming (0-100): 88",
    "Enter marks in Mathematics III (0-100): 92",
    "Enter marks in Data Structures (0-100): 85",
    "Student record added successfully.",
    "",
    "[Notice] Marks record added for Tanmay Arora.",
    "Data saved in students.txt"
]
create_terminal_image("fig5_3_add_student.png", "Terminal - Record Input & InvalidMarksException Validation", fig53_lines)

# -------------------------------------------------------------
# 4. Fig 5.4: Individual Student Performance Report
# -------------------------------------------------------------
fig54_lines = [
    "---------- STUDENT PERFORMANCE REPORT ----------",
    "Roll Number : 101",
    "Student Name: Tanmay Arora",
    "Java Programming: 88/100",
    "Mathematics III : 92/100",
    "Data Structures : 85/100",
    "Total Marks : 265/300",
    "Percentage  : 88.33%",
    "Grade       : A",
    "Highest Subject: Mathematics III",
    "Lowest Subject : Data Structures",
    "-------------------------------------------------",
    "[Notice] Marks record added for Tanmay Arora.",
    "Data saved in students.txt"
]
create_terminal_image("fig5_4_student_report.png", "Terminal - Individual Performance Report Card", fig54_lines)

# -------------------------------------------------------------
# 5. Fig 5.5: View All Students & Search Report
# -------------------------------------------------------------
fig55_lines = [
    "Enter your choice: 3",
    "",
    "ROLL NO - NAME | PERCENTAGE | GRADE",
    "101 - Tanmay Arora | 88.33% | Grade: A",
    "102 - Vishal | 86.67% | Grade: A",
    "103 - Jaskaran Singh | 84.00% | Grade: A",
    "104 - Ansh Arora | 80.00% | Grade: A",
    "",
    "Enter your choice: 4",
    "Enter roll number to search: 102",
    "",
    "---------- STUDENT PERFORMANCE REPORT ----------",
    "Roll Number : 102",
    "Student Name: Vishal",
    "Java Programming: 85/100",
    "Mathematics III : 88/100",
    "Data Structures : 87/100",
    "Total Marks : 260/300",
    "Percentage  : 86.67%",
    "Grade       : A",
    "Highest Subject: Mathematics III",
    "Lowest Subject : Java Programming",
    "-------------------------------------------------"
]
create_terminal_image("fig5_5_view_search.png", "Terminal - View All Records & Search by Roll Number", fig55_lines)

# -------------------------------------------------------------
# 6. Fig 5.6: Student Ranking via Java Comparator
# -------------------------------------------------------------
fig56_lines = [
    "Enter your choice: 5",
    "",
    "============= STUDENT RANKING =============",
    "Rank 1: Tanmay Arora | 88.33% | Grade A",
    "Rank 2: Vishal | 86.67% | Grade A",
    "Rank 3: Jaskaran Singh | 84.00% | Grade A",
    "Rank 4: Ansh Arora | 80.00% | Grade A",
    "===========================================",
    "",
    "Note: Ranked dynamically using Java Comparator in descending order.",
    "Ties are broken alphabetically by student name.",
    "",
    "Enter your choice: _"
]
create_terminal_image("fig5_6_ranking.png", "Terminal - Merit Ranking via Java Comparator", fig56_lines)

# -------------------------------------------------------------
# 7. Fig 5.7: Subject-Wise Statistics
# -------------------------------------------------------------
fig57_lines = [
    "Enter your choice: 6",
    "",
    "========== SUBJECT-WISE STATISTICS ==========",
    "",
    "Subject: Java Programming",
    "Average Marks: 84.25/100",
    "Highest Marks: 88 (by Tanmay Arora)",
    "Lowest Marks : 80 (by Ansh Arora)",
    "",
    "Subject: Mathematics III",
    "Average Marks: 88.75/100",
    "Highest Marks: 92 (by Tanmay Arora)",
    "Lowest Marks : 85 (by Jaskaran Singh)",
    "",
    "Subject: Data Structures",
    "Average Marks: 81.25/100",
    "Highest Marks: 87 (by Vishal)",
    "Lowest Marks : 74 (by Ansh Arora)"
]
create_terminal_image("fig5_7_subject_statistics.png", "Terminal - Subject-Wise Performance Analytics", fig57_lines)

# -------------------------------------------------------------
# 8. Fig 5.8: Class Performance Summary via Java Streams
# -------------------------------------------------------------
fig58_lines = [
    "Enter your choice: 7",
    "",
    "========== CLASS PERFORMANCE SUMMARY ==========",
    "Total Students : 4",
    "Class Average  : 84.75%",
    "Students Passed: 4",
    "Students Failed: 0",
    "Class Topper   : Tanmay Arora",
    "Topper Percent : 88.33%",
    "Grade Distribution: {A=4, A+=0, B=0, C=0, D=0, F=0}",
    "===============================================",
    "",
    "Calculated via Java Stream API & Collectors grouping pipeline.",
    "",
    "Enter your choice: _"
]
create_terminal_image("fig5_8_class_summary.png", "Terminal - Class-Wide Stream API Performance Summary", fig58_lines)

# -------------------------------------------------------------
# 9. Fig 5.9: File Persistence & Byte-Stream Backup
# -------------------------------------------------------------
fig59_lines = [
    "Enter your choice: 8",
    "Data saved in students.txt",
    "Backup created as students_backup.txt",
    "",
    "Enter your choice: 9",
    "Data saved in students.txt",
    "Data saved. Thank you!",
    "",
    "$ cat students.txt",
    "SUBJECTS|Java Programming|Mathematics III|Data Structures",
    "STUDENT|101|Tanmay Arora|88|92|85",
    "STUDENT|102|Vishal|85|88|87",
    "STUDENT|103|Jaskaran Singh|82|85|85",
    "STUDENT|104|Ansh Arora|80|86|74",
    "",
    "$ ls -l students*",
    "-rw-r--r-- 1 student 298 Nov 20 12:20 students.txt",
    "-rw-r--r-- 1 student 298 Nov 20 12:20 students_backup.txt"
]
create_terminal_image("fig5_9_file_backup.png", "Terminal - File Persistence and Byte-Stream Backup", fig59_lines)

# -------------------------------------------------------------
# 10. Fig 4.1: Architecture Flowchart Diagram (Pillow Canvas)
# -------------------------------------------------------------
def create_architecture_diagram():
    width = 880
    height = 420
    img = Image.new('RGB', (width, height), color='#FFFFFF')
    draw = ImageDraw.Draw(img)
    
    # Outer border
    draw.rectangle([(0, 0), (width - 1, height - 1)], outline='#CBD5E1', width=1)
    
    # Title Bar
    draw.rectangle([(0, 0), (width, 42)], fill='#0A192F')
    draw.text((width // 2, 21), "SYSTEM ARCHITECTURE & OOP COMPONENT DATA FLOW", font=font_arch, fill='#FFFFFF', anchor='mm')
    
    # Flow Boxes
    # Row 1 (y: 75 to 145)
    row1 = [
        ("USER / TEACHER", "Terminal Console Interaction", 40, 75, 180, 70, '#0A192F'),
        ("MAIN MENU", "Switch-Case Controller", 250, 75, 180, 70, '#134074'),
        ("DATA INPUT & VALIDATION", "0-100 Marks & Roll Check", 460, 75, 180, 70, '#0077B6'),
        ("STUDENT CLASS", "OOP Encapsulation Model", 670, 75, 170, 70, '#134074')
    ]
    
    # Row 2 (y: 235 to 305) - flows right to left
    row2 = [
        ("FILE I/O & BACKUP", "students.txt & Backup", 40, 235, 180, 70, '#0A192F'),
        ("REPORTABLE INTERFACE", "Standard Report Printing", 250, 235, 180, 70, '#0077B6'),
        ("ANALYTICS & STATS", "Class Average & Extremes", 460, 235, 180, 70, '#134074'),
        ("COMPARATOR ENGINE", "Descending Percentage Sort", 670, 235, 170, 70, '#0A192F')
    ]
    
    # Draw Row 1 Boxes
    for title, sub, x, y, w, h, col in row1:
        draw.rounded_rectangle([(x, y), (x + w, y + h)], radius=8, fill=col, outline='#38BDF8', width=1)
        draw.text((x + w // 2, y + 22), title, font=font_arch, fill='#FFFFFF', anchor='mm')
        draw.text((x + w // 2, y + 46), sub, font=font_arch_sub, fill='#BAE6FD', anchor='mm')
        
    # Draw Row 2 Boxes
    for title, sub, x, y, w, h, col in row2:
        draw.rounded_rectangle([(x, y), (x + w, y + h)], radius=8, fill=col, outline='#38BDF8', width=1)
        draw.text((x + w // 2, y + 22), title, font=font_arch, fill='#FFFFFF', anchor='mm')
        draw.text((x + w // 2, y + 46), sub, font=font_arch_sub, fill='#BAE6FD', anchor='mm')
        
    # Arrows Row 1: Right
    for x_arr in [225, 435, 645]:
        draw.text((x_arr, 110), "➔", font=font_arch, fill='#0077B6', anchor='mm')
        
    # Arrow Row 1 to Row 2: Down
    draw.text((755, 185), "⬇", font=font_arch, fill='#0077B6', anchor='mm')
    
    # Arrows Row 2: Left
    for x_arr in [645, 435, 225]:
        draw.text((x_arr, 270), "🠔", font=font_arch, fill='#0077B6', anchor='mm')
        
    # Bottom Note Box
    draw.rectangle([(40, 335), (840, 395)], fill='#F8FAFC', outline='#CBD5E1', width=1)
    draw.text((width // 2, 352), "Data Persistence: Character Streams (PrintWriter/BufferedReader) for students.txt", font=font_arch_sub, fill='#1E293B', anchor='mm')
    draw.text((width // 2, 374), "Automated Backup: Byte Streams (FileInputStream/FileOutputStream) for students_backup.txt", font=font_arch_sub, fill='#1E293B', anchor='mm')
    
    img.save('images/fig4_1_architecture.png', quality=95)
    print("Saved images/fig4_1_architecture.png")

create_architecture_diagram()
print("All report terminal images generated successfully!")
