const fs = require('fs');
const path = require('path');
const PptxGenJS = require(process.env.APPDATA + '/npm/node_modules/pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9'; // 10 x 5.625 inches
pptx.author = 'Tanmay Arora, Vishal, Jaskaran Singh, Ansh Arora';
pptx.company = 'Chitkara University';
pptx.title = 'Student Performance Analytics Tool - Presentation';

// Academic Palette: Blue, White, Black / Navy
const NAVY = '0A192F';       // Deep Navy
const ROYAL = '134074';      // Slate / Royal Blue
const ACCENT = '0077B6';     // Academic Accent Blue
const SKY = '38BDF8';        // Sky / Light Blue Accent
const LIGHT_BG = 'F8FAFC';   // Clean White-Blue Tint Background
const CARD_BG = 'FFFFFF';    // Card White
const DARK_TEXT = '0F172A';  // Near Black
const BODY_TEXT = '334155';  // Charcoal / Readable Slate
const MUTED_TEXT = '64748B'; // Muted Grey
const BORDER_COLOR = 'CBD5E1';
const GREEN = '166534';
const GREEN_BG = 'DCFCE7';
const RED = '991B1B';
const RED_BG = 'FEE2E2';
const WHITE = 'FFFFFF';

// Helper function to add standard slide header and footer
function addHeader(slide, titleText, subtitleText, slideNumber) {
    // Top banner
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 10, h: 0.82,
        fill: { color: NAVY }
    });
    // Accent line
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0.82, w: 10, h: 0.05,
        fill: { color: ACCENT }
    });
    // Title
    slide.addText(titleText, {
        x: 0.6, y: 0.1, w: 8.0, h: 0.4,
        fontSize: 18, fontFace: 'Arial', bold: true, color: WHITE
    });
    // Subtitle
    if (subtitleText) {
        slide.addText(subtitleText, {
            x: 0.6, y: 0.48, w: 8.0, h: 0.28,
            fontSize: 10.5, fontFace: 'Arial', color: '93C5FD'
        });
    }
    // Slide Number badge
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 8.85, y: 0.18, w: 0.65, h: 0.44,
        fill: { color: ROYAL },
        line: { color: SKY, width: 1 },
        rectRadius: 0.06
    });
    slide.addText(slideNumber.toString(), {
        x: 8.85, y: 0.18, w: 0.65, h: 0.44,
        fontSize: 13, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle'
    });

    // Bottom Footer
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 5.35, w: 10, h: 0.275,
        fill: { color: 'F1F5F9' },
        line: { color: BORDER_COLOR, width: 0.5 }
    });
    slide.addText('B.Tech CSE (AI & ML) | 3rd Semester | Chitkara University', {
        x: 0.6, y: 5.38, w: 5.0, h: 0.22,
        fontSize: 8.5, fontFace: 'Arial', color: MUTED_TEXT
    });
    slide.addText('Student Performance Analytics Tool | Java Programming', {
        x: 5.4, y: 5.38, w: 4.0, h: 0.22,
        fontSize: 8.5, fontFace: 'Arial', color: ROYAL, align: 'right', bold: true
    });
}

// =============================================================
// SLIDE 1: TITLE PAGE
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: NAVY };

    // Decorative vertical bars on left
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.25, h: 5.625, fill: { color: ACCENT } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.25, y: 0, w: 0.08, h: 5.625, fill: { color: SKY } });

    // Top Category Badge
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 0.45, w: 5.4, h: 0.38,
        fill: { color: ROYAL }, line: { color: SKY, width: 1 }, rectRadius: 0.08
    });
    slide.addText('B.TECH CSE (AI & ML) — 3RD SEMESTER COLLEGE PROJECT', {
        x: 0.8, y: 0.45, w: 5.4, h: 0.38,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: 'BAE6FD', align: 'center', valign: 'middle'
    });

    // Main Title
    slide.addText('STUDENT PERFORMANCE\nANALYTICS TOOL', {
        x: 0.8, y: 0.98, w: 8.5, h: 1.25,
        fontSize: 27, fontFace: 'Arial', bold: true, color: WHITE, lineSpacing: 34
    });

    // Subtitle
    slide.addText('A Console-Based Core Java Application for Academic Record Management & Analysis', {
        x: 0.8, y: 2.32, w: 8.5, h: 0.35,
        fontSize: 12, fontFace: 'Arial', color: '93C5FD'
    });

    // Accent line
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.75, w: 8.4, h: 0.02, fill: { color: SKY } });

    // Left Card: Academic Details
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 2.95, w: 4.0, h: 2.25,
        fill: { color: '0F172A' }, line: { color: '334155', width: 1 }, rectRadius: 0.08
    });
    slide.addText('ACADEMIC DETAILS', {
        x: 1.0, y: 3.1, w: 3.6, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: SKY
    });
    const acadDetails = [
        'Course: B.Tech CSE (AI & ML)',
        'Semester: 3rd Semester',
        'Subject: Java Programming',
        'Institution: Chitkara University',
        'Interface: Command Line / Terminal'
    ];
    slide.addText(acadDetails.join('\n'), {
        x: 1.0, y: 3.42, w: 3.6, h: 1.6,
        fontSize: 10, fontFace: 'Arial', color: WHITE, lineSpacing: 18
    });

    // Right Card: Project Team Members (Exact Order!)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 2.95, w: 4.1, h: 2.25,
        fill: { color: '0F172A' }, line: { color: '334155', width: 1 }, rectRadius: 0.08
    });
    slide.addText('PROJECT TEAM MEMBERS', {
        x: 5.3, y: 3.1, w: 3.7, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: SKY
    });
    const teamDetails = [
        '1. Vishal           (Roll No: 2510992990)',
        '2. Jaskaran Singh   (Roll No: 2510993488)',
        '3. Ansh Arora       (Roll No: 2510992560)',
        '4. Tanmay Arora     (Roll No: 2510993577)'
    ];
    slide.addText(teamDetails.join('\n'), {
        x: 5.3, y: 3.42, w: 3.7, h: 1.6,
        fontSize: 10.5, fontFace: 'Arial', color: WHITE, lineSpacing: 20
    });
}

// =============================================================
// SLIDE 2: INTRODUCTION
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Introduction', 'What is the Student Performance Analytics Tool?', 2);

    // Top Summary Banner
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 8.8, h: 0.75,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Our project is a console-based Student Performance Analytics Tool developed using Core Java. It helps teachers easily manage student marks, calculate grades automatically, and analyze class performance without any manual errors.', {
        x: 0.8, y: 1.1, w: 8.4, h: 0.58,
        fontSize: 11, fontFace: 'Arial', color: DARK_TEXT, lineSpacing: 16
    });

    // 4 Key Capability Cards (2x2 Grid)
    const cards = [
        {
            title: '1. Student Data Management',
            points: [
                'Developed entirely using Core Java programming.',
                'Allows teachers to store student roll numbers and names.',
                'Supports setting up marks for multiple custom subjects.'
            ],
            x: 0.6, y: 1.95
        },
        {
            title: '2. Automatic Percentage & Grades',
            points: [
                'Automatically calculates total marks and percentages.',
                'Assigns academic letter grades from A+ down to F.',
                'Eliminates manual calculation mistakes completely.'
            ],
            x: 5.1, y: 1.95
        },
        {
            title: '3. Student Ranking System',
            points: [
                'Compares overall percentages of all students in the class.',
                'Ranks students automatically from highest to lowest score.',
                'Uses Java Comparator for fair and accurate sorting.'
            ],
            x: 0.6, y: 3.55
        },
        {
            title: '4. Performance Analytics & Reports',
            points: [
                'Identifies the highest and lowest scoring subjects for each student.',
                'Generates subject averages, top scorers, and lowest scores.',
                'Produces class performance summaries with grade distributions.'
            ],
            x: 5.1, y: 3.55
        }
    ];

    cards.forEach(card => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: card.x, y: card.y, w: 4.3, h: 1.45,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: card.x, y: card.y, w: 4.3, h: 0.32,
            fill: { color: ROYAL }
        });
        slide.addText(card.title, {
            x: card.x + 0.15, y: card.y + 0.04, w: 4.0, h: 0.25,
            fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(card.points.map(p => '•  ' + p).join('\n'), {
            x: card.x + 0.15, y: card.y + 0.38, w: 4.0, h: 0.98,
            fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
        });
    });
}

// =============================================================
// SLIDE 3: PROBLEM STATEMENT
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Problem Statement', 'Challenges of Manual Student Record Management & Our Solution', 3);

    // Left Column: The Problem (Manual Process)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 4.25, h: 4.1,
        fill: { color: CARD_BG }, line: { color: 'FCA5A5', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 4.25, h: 0.45,
        fill: { color: RED }
    });
    slide.addText('CHALLENGES OF MANUAL MANAGEMENT', {
        x: 0.75, y: 1.15, w: 3.9, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const problems = [
        {
            head: '1. Time-Consuming Calculations',
            desc: 'Calculating totals and percentages manually for 50+ students takes significant teacher time.'
        },
        {
            head: '2. Human Calculation Errors',
            desc: 'Manual ranking and percentage calculations often lead to mistakes and grading discrepancies.'
        },
        {
            head: '3. Difficulty Comparing Subjects',
            desc: 'Comparing scores across multiple subjects on paper or spreadsheets is tedious and confusing.'
        },
        {
            head: '4. Disorganized Student Records',
            desc: 'Physical mark registers can become messy, damaged, or misplaced over an academic semester.'
        },
        {
            head: '5. Difficult Class Reports',
            desc: 'Finding the class average, topper, and grade counts requires extra manual effort and formulas.'
        }
    ];

    let probY = 1.6;
    problems.forEach(item => {
        slide.addText('• ' + item.head, {
            x: 0.8, y: probY, w: 3.85, h: 0.22,
            fontSize: 10, fontFace: 'Arial', bold: true, color: RED
        });
        slide.addText(item.desc, {
            x: 0.95, y: probY + 0.22, w: 3.7, h: 0.35,
            fontSize: 9, fontFace: 'Arial', color: BODY_TEXT
        });
        probY += 0.64;
    });

    // Right Column: Our Java Solution
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 1.05, w: 4.25, h: 4.1,
        fill: { color: CARD_BG }, line: { color: '86EFAC', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.15, y: 1.05, w: 4.25, h: 0.45,
        fill: { color: GREEN }
    });
    slide.addText('OUR AUTOMATED JAVA SOLUTION', {
        x: 5.3, y: 1.15, w: 3.9, h: 0.3,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const solutions = [
        {
            head: '1. Instant Automatic Calculations',
            desc: 'Percentages and total marks are computed instantly upon mark entry with zero calculation errors.'
        },
        {
            head: '2. Automated Merit Ranking',
            desc: 'Java Comparator sorts all students automatically based on percentage, breaking ties by name.'
        },
        {
            head: '3. Comprehensive Analytics',
            desc: 'Immediately displays subject-wise averages, top scorers, and lowest scores in clean tables.'
        },
        {
            head: '4. Permanent File Storage',
            desc: 'Records are saved in students.txt with an automatic byte-stream backup in students_backup.txt.'
        },
        {
            head: '5. Fast Console Navigation',
            desc: 'A user-friendly numbered menu lets teachers execute all tasks in seconds with simple inputs.'
        }
    ];

    let solY = 1.6;
    solutions.forEach(item => {
        slide.addText('✔ ' + item.head, {
            x: 5.35, y: solY, w: 3.85, h: 0.22,
            fontSize: 10, fontFace: 'Arial', bold: true, color: GREEN
        });
        slide.addText(item.desc, {
            x: 5.5, y: solY + 0.22, w: 3.7, h: 0.35,
            fontSize: 9, fontFace: 'Arial', color: BODY_TEXT
        });
        solY += 0.64;
    });
}

// =============================================================
// SLIDE 4: OBJECTIVES OF THE PROJECT
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Project Objectives', 'Core Goals and Capabilities of the System', 4);

    const objectives = [
        { num: '01', title: 'Multi-Subject Marks Entry', desc: 'Accept and configure marks for any number of custom academic subjects dynamically.' },
        { num: '02', title: 'Automatic Percentage Calculation', desc: 'Calculate student percentages automatically using standard mathematical formulas.' },
        { num: '03', title: 'Accurate Letter Grading', desc: 'Assign standard letter grades (A+, A, B, C, D, F) based on percentage thresholds.' },
        { num: '04', title: 'Subject Strengths & Weaknesses', desc: 'Identify the highest and lowest scoring subjects for every individual student.' },
        { num: '05', title: 'Automated Student Ranking', desc: 'Rank students fairly from highest to lowest percentage using Java Comparator.' },
        { num: '06', title: 'Class-Wide Performance Tracking', desc: 'Calculate overall class average, passing/failing rates, and identify the class topper.' },
        { num: '07', title: 'Subject-Wise Statistical Analysis', desc: 'Generate subject statistics including class average, top scorer, and lowest score.' },
        { num: '08', title: 'Persistent Storage & Backup', desc: 'Save all student records safely in text files and generate automated backup copies.' }
    ];

    objectives.forEach((obj, idx) => {
        const col = idx < 4 ? 0 : 1;
        const row = idx % 4;
        const x = col === 0 ? 0.6 : 5.15;
        const y = 1.05 + row * 1.02;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: y, w: 4.25, h: 0.92,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });

        // Number Badge
        slide.addShape(pptx.ShapeType.roundRect, {
            x: x + 0.12, y: y + 0.15, w: 0.5, h: 0.6,
            fill: { color: ROYAL }, rectRadius: 0.06
        });
        slide.addText(obj.num, {
            x: x + 0.12, y: y + 0.15, w: 0.5, h: 0.6,
            fontSize: 13, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle'
        });

        // Content
        slide.addText(obj.title, {
            x: x + 0.72, y: y + 0.12, w: 3.4, h: 0.25,
            fontSize: 10.5, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(obj.desc, {
            x: x + 0.72, y: y + 0.38, w: 3.4, h: 0.48,
            fontSize: 9, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 13
        });
    });
}

// =============================================================
// SLIDE 5: TECHNOLOGIES AND JAVA CONCEPTS USED
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Technologies & Java Concepts Used', 'Core Java Tools and Object-Oriented Principles Implemented', 5);

    // Left Column: General Tech Stack Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 3.0, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 3.0, h: 0.38,
        fill: { color: NAVY }
    });
    slide.addText('DEVELOPMENT STACK', {
        x: 0.75, y: 1.12, w: 2.7, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const techStack = [
        { label: 'Language', val: 'Java (JDK 8+)' },
        { label: 'Paradigm', val: 'Object-Oriented (OOP)' },
        { label: 'Interface', val: 'Terminal / Command Line' },
        { label: 'Data Storage', val: 'Character Stream (students.txt)' },
        { label: 'Backup Engine', val: 'Byte Stream (students_backup.txt)' },
        { label: 'IDE Used', val: 'VS Code / IntelliJ IDEA' },
        { label: 'Version Control', val: 'Git & GitHub' }
    ];

    let stackY = 1.55;
    techStack.forEach(item => {
        slide.addText(item.label, {
            x: 0.75, y: stackY, w: 2.7, h: 0.18,
            fontSize: 9, fontFace: 'Arial', bold: true, color: ACCENT
        });
        slide.addText(item.val, {
            x: 0.75, y: stackY + 0.18, w: 2.7, h: 0.26,
            fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT
        });
        stackY += 0.48;
    });

    // Right Column: Java Concepts Table
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 3.8, y: 1.05, w: 5.6, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 3.8, y: 1.05, w: 5.6, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('JAVA CONCEPT', {
        x: 3.95, y: 1.12, w: 2.0, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });
    slide.addText('HOW IT IS USED IN OUR PROJECT', {
        x: 5.95, y: 1.12, w: 3.3, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const concepts = [
        { c: 'Classes & Objects', u: 'To create and manage real-world student records.' },
        { c: 'Encapsulation', u: 'To protect student information using private variables and public getters.' },
        { c: 'ArrayList', u: 'To store dynamic lists of students and configured subjects.' },
        { c: 'HashMap', u: 'To store marks of different subjects as key-value pairs.' },
        { c: 'Interface (Reportable)', u: 'To define a common contract for printing student report cards.' },
        { c: 'Custom Exception', u: 'InvalidMarksException validates marks strictly between 0 and 100.' },
        { c: 'Comparator', u: 'To rank students according to percentage and sort ties by name.' },
        { c: 'Stream API', u: 'To calculate class averages, pass/fail counts, and grade distributions.' },
        { c: 'Generics', u: 'To display lists of different data types using one displayList() method.' },
        { c: 'Multithreading', u: 'To run a background notification thread when records are added.' },
        { c: 'File Handling', u: 'To save records in students.txt and create students_backup.txt.' }
    ];

    let rowY = 1.48;
    concepts.forEach((row, i) => {
        const bg = (i % 2 === 0) ? 'FFFFFF' : 'F8FAFC';
        slide.addShape(pptx.ShapeType.rect, {
            x: 3.85, y: rowY, w: 5.5, h: 0.31,
            fill: { color: bg }
        });
        slide.addText(row.c, {
            x: 3.95, y: rowY + 0.04, w: 1.95, h: 0.24,
            fontSize: 8.5, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(row.u, {
            x: 5.95, y: rowY + 0.04, w: 3.35, h: 0.24,
            fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT
        });
        rowY += 0.32;
    });
}

// =============================================================
// SLIDE 6: PROJECT ARCHITECTURE
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Project Architecture', 'System Workflow from User Input to File Persistence', 6);

    // Architecture Flow Steps (8 boxes connected by arrows)
    const flowSteps = [
        { label: 'USER\n(Teacher / Evaluator)', color: NAVY, textCol: WHITE },
        { label: 'MAIN MENU\n(Switch-Case Options)', color: ROYAL, textCol: WHITE },
        { label: 'STUDENT DATA INPUT\n(Roll, Name, Marks)', color: ACCENT, textCol: WHITE },
        { label: 'STUDENT CLASS\n(OOP Encapsulation)', color: ROYAL, textCol: WHITE },
        { label: 'PERFORMANCE ENGINE\n(Total, % & Grades)', color: NAVY, textCol: WHITE },
        { label: 'RANKING & ANALYTICS\n(Comparator & Streams)', color: ROYAL, textCol: WHITE },
        { label: 'REPORT GENERATION\n(Reportable Interface)', color: ACCENT, textCol: WHITE },
        { label: 'FILE PERSISTENCE\n(students.txt & Backup)', color: NAVY, textCol: WHITE }
    ];

    // Arrange in 2 rows of 4 boxes with connecting arrows
    flowSteps.forEach((step, idx) => {
        const isRow1 = idx < 4;
        const col = isRow1 ? idx : (7 - idx); // row 2 goes right to left for a clean snake flow
        const x = 0.6 + col * 2.25;
        const y = isRow1 ? 1.15 : 2.75;

        // Step Box
        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: y, w: 2.05, h: 1.0,
            fill: { color: step.color },
            line: { color: SKY, width: 1 },
            rectRadius: 0.08
        });
        slide.addText(step.label, {
            x: x + 0.08, y: y + 0.12, w: 1.89, h: 0.76,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: step.textCol, align: 'center', valign: 'middle'
        });

        // Connector Arrow
        if (idx < 3) {
            // Right arrow in row 1
            slide.addText('➔', {
                x: x + 2.05, y: y + 0.3, w: 0.25, h: 0.35,
                fontSize: 14, color: ACCENT, align: 'center'
            });
        } else if (idx === 3) {
            // Down arrow from row 1 to row 2
            slide.addText('⬇', {
                x: x + 0.85, y: y + 1.05, w: 0.35, h: 0.45,
                fontSize: 16, color: ACCENT, align: 'center'
            });
        } else if (idx >= 4 && idx < 7) {
            // Left arrow in row 2
            slide.addText('➔', {
                x: x - 0.25, y: y + 0.3, w: 0.25, h: 0.35,
                fontSize: 14, color: ACCENT, align: 'center'
            });
        }
    });

    // Bottom Architecture Explanation Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 4.05, w: 8.8, h: 1.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 4.05, w: 8.8, h: 0.3,
        fill: { color: ROYAL }
    });
    slide.addText('HOW THE ARCHITECTURE OPERATES (EASY TO EXPLAIN)', {
        x: 0.75, y: 4.1, w: 8.5, h: 0.22,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE
    });
    const archExplanation = [
        '1. The user interacts through a simple Main Menu that directs requests to specific functions.',
        '2. Student inputs are encapsulated safely inside Student objects and stored in memory using ArrayList and HashMap.',
        '3. The Calculation & Analytics engines compute percentages, grades, rankings, and class-wide statistics.',
        '4. All data is saved permanently to students.txt and backed up into students_backup.txt using Java File I/O.'
    ];
    slide.addText(archExplanation.join('\n'), {
        x: 0.75, y: 4.4, w: 8.5, h: 0.68,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 13
    });
}

// =============================================================
// SLIDE 7: STUDENT REGISTRATION AND DATA MANAGEMENT
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Student Registration & Data Management', 'How Student Information is Collected, Validated, and Stored', 7);

    // 3 Cards across the slide
    const regCards = [
        {
            title: '1. Information Collected',
            badge: 'INPUT DATA',
            badgeCol: ROYAL,
            points: [
                'Student Roll Number: Unique numerical identifier for each student.',
                'Student Name: Cleaned string identifying the student.',
                'Subject Names: Defined once during subject setup (e.g. Java, Math, DSA).',
                'Subject Marks: Score obtained in each subject (0 to 100).'
            ],
            x: 0.6, w: 2.8
        },
        {
            title: '2. Input Validation Rules',
            badge: 'DATA SAFETY',
            badgeCol: GREEN,
            points: [
                'Duplicate Roll Prevention: The program searches existing records to ensure roll numbers are never repeated.',
                'Marks Range Check (0–100): Uses custom InvalidMarksException to reject negative numbers or marks > 100.',
                'Safe Loop Handling: Keeps re-asking user until valid numbers are provided.'
            ],
            x: 3.6, w: 2.8
        },
        {
            title: '3. Data Structures Used',
            badge: 'JAVA COLLECTIONS',
            badgeCol: ACCENT,
            points: [
                'ArrayList<Student>: Dynamically stores all student objects in memory without size limits.',
                'HashMap<String, Integer>: Stores subject names as keys and marks as values for each student.',
                'ArrayList<String>: Holds the list of registered academic subjects.'
            ],
            x: 6.6, w: 2.8
        }
    ];

    regCards.forEach(card => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: card.x, y: 1.05, w: card.w, h: 2.9,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: card.x, y: 1.05, w: card.w, h: 0.38,
            fill: { color: card.badgeCol }
        });
        slide.addText(card.title, {
            x: card.x + 0.1, y: 1.12, w: card.w - 0.2, h: 0.25,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(card.points.map(p => '•  ' + p).join('\n\n'), {
            x: card.x + 0.12, y: 1.5, w: card.w - 0.24, h: 2.35,
            fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
        });
    });

    // Bottom Card: Code Flow Explanation
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 4.1, w: 8.8, h: 1.05,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Key Takeaway for Viva:', {
        x: 0.8, y: 4.2, w: 3.0, h: 0.22,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: NAVY
    });
    slide.addText('"The program strictly prevents bad data. If a user enters a roll number that already exists, it warns the user immediately. If a mark is entered outside 0 to 100, InvalidMarksException is caught and the user is prompted again without crashing the application."', {
        x: 0.8, y: 4.45, w: 8.4, h: 0.6,
        fontSize: 9, fontFace: 'Arial', color: DARK_TEXT, italic: true, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 8: PERCENTAGE AND GRADE CALCULATION
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Percentage & Grade Calculation', 'Formulas, Step-by-Step Example, and Academic Grading System', 8);

    // Left Column: Formula & Example Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 4.3, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('PERCENTAGE CALCULATION & EXAMPLE', {
        x: 0.75, y: 1.13, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    // Formula Box
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.55, w: 3.9, h: 0.75,
        fill: { color: 'F0F7FF' }, line: { color: 'BAE6FD', width: 1 }, rectRadius: 0.06
    });
    slide.addText('MATHEMATICAL FORMULA:', {
        x: 0.9, y: 1.62, w: 3.7, h: 0.2,
        fontSize: 8.5, fontFace: 'Arial', bold: true, color: ACCENT
    });
    slide.addText('Percentage = (Total Marks Obtained / Maximum Marks) × 100', {
        x: 0.9, y: 1.88, w: 3.7, h: 0.35,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: NAVY
    });

    // Step-by-Step Example
    slide.addText('PRACTICAL EXAMPLE (3 SUBJECTS):', {
        x: 0.8, y: 2.4, w: 3.9, h: 0.22,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: DARK_TEXT
    });
    const exampleLines = [
        '• Subject 1: Java Programming   = 85 / 100',
        '• Subject 2: Mathematics        = 90 / 100',
        '• Subject 3: Data Structures    = 75 / 100',
        '--------------------------------------------',
        'Total Marks Obtained  = 85 + 90 + 75 = 250',
        'Maximum Total Marks   = 3 × 100 = 300',
        'Percentage = (250 / 300) × 100 = 83.33%',
        'Calculated Grade = A (Excellent)'
    ];
    slide.addText(exampleLines.join('\n'), {
        x: 0.8, y: 2.68, w: 3.9, h: 2.3,
        fontSize: 9, fontFace: 'Courier New', color: BODY_TEXT, lineSpacing: 15
    });

    // Right Column: Grading System Table
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.05, w: 4.3, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.05, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('OFFICIAL ACADEMIC GRADING SCALE', {
        x: 5.25, y: 1.13, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    // Table Header
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.2, y: 1.55, w: 4.1, h: 0.32,
        fill: { color: 'E2E8F0' }
    });
    slide.addText('PERCENTAGE RANGE', { x: 5.3, y: 1.62, w: 1.6, h: 0.2, fontSize: 8.5, fontFace: 'Arial', bold: true, color: NAVY });
    slide.addText('GRADE', { x: 7.0, y: 1.62, w: 0.8, h: 0.2, fontSize: 8.5, fontFace: 'Arial', bold: true, color: NAVY });
    slide.addText('EVALUATION', { x: 7.9, y: 1.62, w: 1.3, h: 0.2, fontSize: 8.5, fontFace: 'Arial', bold: true, color: NAVY });

    const grades = [
        { range: '90% – 100%', grade: 'A+', remark: 'Outstanding', bg: 'FFFFFF', color: GREEN },
        { range: '80% – 89%', grade: 'A', remark: 'Excellent', bg: 'F8FAFC', color: GREEN },
        { range: '70% – 79%', grade: 'B', remark: 'Very Good', bg: 'FFFFFF', color: NAVY },
        { range: '60% – 69%', grade: 'C', remark: 'Good', bg: 'F8FAFC', color: NAVY },
        { range: '50% – 59%', grade: 'D', remark: 'Pass', bg: 'FFFFFF', color: 'B45309' },
        { range: 'Below 50%', grade: 'F', remark: 'Fail / Reappear', bg: 'FEE2E2', color: RED }
    ];

    let gY = 1.9;
    grades.forEach(g => {
        slide.addShape(pptx.ShapeType.rect, {
            x: 5.2, y: gY, w: 4.1, h: 0.35,
            fill: { color: g.bg }
        });
        slide.addText(g.range, { x: 5.3, y: gY + 0.08, w: 1.6, h: 0.2, fontSize: 8.5, fontFace: 'Arial', color: DARK_TEXT });
        slide.addText(g.grade, { x: 7.0, y: gY + 0.08, w: 0.8, h: 0.2, fontSize: 9.5, fontFace: 'Arial', bold: true, color: g.color });
        slide.addText(g.remark, { x: 7.9, y: gY + 0.08, w: 1.3, h: 0.2, fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT });
        gY += 0.37;
    });

    slide.addText('Implementation Note:\nThe getGrade() method uses simple if-else conditional branches to assign grades instantly based on the percentage calculation.', {
        x: 5.3, y: 4.35, w: 3.9, h: 0.65,
        fontSize: 8.5, fontFace: 'Arial', color: MUTED_TEXT, lineSpacing: 13
    });
}

// =============================================================
// SLIDE 9: STUDENT RANKING USING COMPARATOR
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Student Ranking Using Comparator', 'How the Program Sorts Students by Academic Performance', 9);

    // Left Column: How Comparator Works
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 4.3, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('HOW JAVA COMPARATOR WORKS', {
        x: 0.75, y: 1.13, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const rankPoints = [
        '1. Objective Ranking: The program compares percentages of all registered students in the class.',
        '2. Descending Sort: Higher percentages receive higher ranks. The student with the highest percentage gets Rank 1.',
        '3. Java Comparator Interface: Implemented via Collections.sort() using a custom chained comparator.',
        '4. Alphabetical Tie-Breaker: If two students achieve the exact same percentage, their names are sorted alphabetically.',
        '5. Safe In-Memory Copy: Ranking is performed on a cloned ArrayList so the original student registration order is preserved.'
    ];

    slide.addText(rankPoints.join('\n\n'), {
        x: 0.8, y: 1.6, w: 3.9, h: 2.6,
        fontSize: 9, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });

    // Code snippet preview
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 3.8, w: 3.9, h: 1.15,
        fill: { color: '0F172A' }, rectRadius: 0.06
    });
    slide.addText('// Java Implementation in Project:\nComparator<Student> comp = \n    Comparator.comparingDouble(Student::getPercentage)\n    .reversed().thenComparing(Student::getName);\nCollections.sort(rankedStudents, comp);', {
        x: 0.9, y: 3.88, w: 3.7, h: 1.0,
        fontSize: 8, fontFace: 'Courier New', color: '38BDF8', lineSpacing: 12
    });

    // Right Column: Sample Ranking Output
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.05, w: 4.3, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.05, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('SAMPLE RANKING LEADERBOARD', {
        x: 5.25, y: 1.13, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const rankSample = [
        { rank: 'Rank 1', name: 'Tanmay Arora', pct: '92.00%', grade: 'A+', badge: 'TOPPER', bg: 'DCFCE7', col: GREEN },
        { rank: 'Rank 2', name: 'Vishal', pct: '88.33%', grade: 'A', badge: 'EXCELLENT', bg: 'F0F7FF', col: NAVY },
        { rank: 'Rank 3', name: 'Jaskaran Singh', pct: '85.67%', grade: 'A', badge: 'EXCELLENT', bg: 'FFFFFF', col: NAVY },
        { rank: 'Rank 4', name: 'Ansh Arora', pct: '81.00%', grade: 'A', badge: 'EXCELLENT', bg: 'F8FAFC', col: NAVY },
        { rank: 'Rank 5', name: 'Rohan Verma', pct: '74.50%', grade: 'B', badge: 'VERY GOOD', bg: 'FFFFFF', col: BODY_TEXT },
        { rank: 'Rank 6', name: 'Priya Sharma', pct: '62.00%', grade: 'C', badge: 'GOOD', bg: 'F8FAFC', col: BODY_TEXT }
    ];

    let rY = 1.55;
    rankSample.forEach(r => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.25, y: rY, w: 4.0, h: 0.38,
            fill: { color: r.bg }, line: { color: BORDER_COLOR, width: 0.5 }, rectRadius: 0.04
        });
        slide.addText(r.rank, { x: 5.35, y: rY + 0.08, w: 0.8, h: 0.22, fontSize: 9, fontFace: 'Arial', bold: true, color: r.col });
        slide.addText(r.name, { x: 6.15, y: rY + 0.08, w: 1.4, h: 0.22, fontSize: 9, fontFace: 'Arial', bold: true, color: DARK_TEXT });
        slide.addText(r.pct + ' (' + r.grade + ')', { x: 7.55, y: rY + 0.08, w: 1.0, h: 0.22, fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT });
        slide.addShape(pptx.ShapeType.rect, { x: 8.55, y: rY + 0.08, w: 0.6, h: 0.22, fill: { color: r.col } });
        slide.addText(r.badge, { x: 8.55, y: rY + 0.08, w: 0.6, h: 0.22, fontSize: 6.5, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle' });
        rY += 0.42;
    });

    slide.addText('Key Takeaway for Viva:\n"Using Java Comparator guarantees that sorting is automated, fair, and free of manual human errors."', {
        x: 5.3, y: 4.25, w: 3.9, h: 0.7,
        fontSize: 8.5, fontFace: 'Arial', color: NAVY, italic: true, lineSpacing: 13
    });
}

// =============================================================
// SLIDE 10: SUBJECT-WISE STATISTICS
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Subject-Wise Statistics', 'Analyzing Subject Difficulty, Average Performance, and Top Scorers', 10);

    // Top Summary Banner
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 8.8, h: 0.65,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Option 6 in our program analyzes every subject across all student records. It calculates the subject average, finds the highest marks scored (with the top student name), and finds the lowest marks scored.', {
        x: 0.8, y: 1.13, w: 8.4, h: 0.48,
        fontSize: 10, fontFace: 'Arial', color: DARK_TEXT, lineSpacing: 15
    });

    // Center Table Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.85, w: 8.8, h: 1.85,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.85, w: 8.8, h: 0.35,
        fill: { color: NAVY }
    });

    // Table Header
    slide.addText('SUBJECT NAME', { x: 0.8, y: 1.92, w: 2.2, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('CLASS AVERAGE', { x: 3.0, y: 1.92, w: 1.8, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('HIGHEST SCORE & TOPPER', { x: 4.8, y: 1.92, w: 2.4, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('LOWEST SCORE', { x: 7.2, y: 1.92, w: 2.0, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: WHITE });

    const subjectStats = [
        { sub: 'Java Programming', avg: '78.50 / 100', high: '92 / 100 (Aman Sharma)', low: '65 / 100 (Priya Verma)', bg: 'FFFFFF' },
        { sub: 'Mathematics III', avg: '81.00 / 100', high: '95 / 100 (Rahul Gupta)', low: '58 / 100 (Simran Kaur)', bg: 'F8FAFC' },
        { sub: 'Data Structures (DSA)', avg: '74.00 / 100', high: '88 / 100 (Neha Patel)', low: '60 / 100 (Rohan Mehta)', bg: 'FFFFFF' }
    ];

    let subY = 2.25;
    subjectStats.forEach(st => {
        slide.addShape(pptx.ShapeType.rect, {
            x: 0.65, y: subY, w: 8.7, h: 0.38,
            fill: { color: st.bg }
        });
        slide.addText(st.sub, { x: 0.8, y: subY + 0.08, w: 2.2, h: 0.22, fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY });
        slide.addText(st.avg, { x: 3.0, y: subY + 0.08, w: 1.8, h: 0.22, fontSize: 9, fontFace: 'Arial', color: BODY_TEXT });
        slide.addText(st.high, { x: 4.8, y: subY + 0.08, w: 2.4, h: 0.22, fontSize: 9, fontFace: 'Arial', bold: true, color: GREEN });
        slide.addText(st.low, { x: 7.2, y: subY + 0.08, w: 2.0, h: 0.22, fontSize: 9, fontFace: 'Arial', color: RED });
        subY += 0.42;
    });

    // 3 Benefit Cards below
    const subBenefits = [
        {
            title: 'Identify Strong Subjects',
            desc: 'Shows subjects where class average is high, showing students understand concepts well.'
        },
        {
            title: 'Detect Difficult Subjects',
            desc: 'Highlights subjects with lower averages so teachers can schedule remedial sessions.'
        },
        {
            title: 'Recognize Individual Talent',
            desc: 'Displays the specific student name who scored the highest in each subject.'
        }
    ];

    subBenefits.forEach((b, idx) => {
        const x = 0.6 + idx * 3.05;
        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: 3.85, w: 2.8, h: 1.3,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addText('✔ ' + b.title, {
            x: x + 0.15, y: 3.95, w: 2.5, h: 0.25,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: ROYAL
        });
        slide.addText(b.desc, {
            x: x + 0.15, y: 4.25, w: 2.5, h: 0.8,
            fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 13
        });
    });
}

// =============================================================
// SLIDE 11: CLASS PERFORMANCE SUMMARY
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Class Performance Summary', 'Batch-Level Academic Insights Powered by Java Stream API', 11);

    // Left Column: Key Metrics Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 4.3, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('OVERALL CLASS PERFORMANCE METRICS', {
        x: 0.75, y: 1.13, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const metrics = [
        { label: 'Total Students Enrolled', val: '5 Students in Sample Batch' },
        { label: 'Overall Class Average', val: '84.31% Mean Score' },
        { label: 'Students Passed (Grade A+ to D)', val: '5 Students (100% Pass Rate)' },
        { label: 'Students Failed (Grade F)', val: '0 Students' },
        { label: 'Class Topper Name', val: 'Tanmay Arora' },
        { label: 'Class Topper Percentage', val: '92.00% (Grade A+)' },
        { label: 'Grade Distribution Map', val: '{A+=1, A=3, B=1, C=0, D=0, F=0}' }
    ];

    let mY = 1.55;
    metrics.forEach(m => {
        slide.addText('• ' + m.label + ':', {
            x: 0.8, y: mY, w: 3.9, h: 0.18,
            fontSize: 8.5, fontFace: 'Arial', bold: true, color: ACCENT
        });
        slide.addText(m.val, {
            x: 0.95, y: mY + 0.18, w: 3.75, h: 0.22,
            fontSize: 9, fontFace: 'Arial', color: DARK_TEXT
        });
        mY += 0.48;
    });

    // Right Column: Java Stream API Role
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.05, w: 4.3, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.05, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('HOW JAVA STREAM API IS USED', {
        x: 5.25, y: 1.13, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const streamExplanations = [
        {
            goal: '1. Class Average Calculation',
            code: 'students.stream().mapToDouble(Student::getPercentage).average()',
            desc: 'Streams map each student to percentage and compute the mean.'
        },
        {
            goal: '2. Finding the Class Topper',
            code: 'students.stream().max(Comparator.comparingDouble(...))',
            desc: 'Finds the student with the maximum percentage without manual loops.'
        },
        {
            goal: '3. Pass / Fail Counts',
            code: 'students.stream().filter(s -> !s.getGrade().equals("F")).count()',
            desc: 'Filters out failing students to count passed candidates.'
        },
        {
            goal: '4. Grade Distribution',
            code: 'Collectors.groupingBy(Student::getGrade, TreeMap::new, ...)',
            desc: 'Groups students by grade letter and counts frequencies in sorted order.'
        }
    ];

    let stY = 1.55;
    streamExplanations.forEach(st => {
        slide.addText(st.goal, {
            x: 5.25, y: stY, w: 3.9, h: 0.2,
            fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(st.desc, {
            x: 5.25, y: stY + 0.2, w: 3.9, h: 0.35,
            fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT
        });
        stY += 0.62;
    });

    slide.addText('Takeaway for Viva:\n"Stream API allows us to process large student collections cleanly with compact, readable code instead of writing multiple nested for-loops."', {
        x: 5.25, y: 4.25, w: 3.9, h: 0.7,
        fontSize: 8.5, fontFace: 'Arial', color: NAVY, italic: true, lineSpacing: 13
    });
}

// =============================================================
// SLIDE 12: FILE HANDLING AND BACKUP SYSTEM
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'File Handling & Backup System', 'Persistent Storage Architecture and Automated Byte-Stream Backup', 12);

    // Left Column: Main Storage (Character Streams)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 4.3, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('DATA STORAGE (CHARACTER STREAMS)', {
        x: 0.75, y: 1.13, w: 4.0, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const charStreamPoints = [
        '• File Used: students.txt',
        '• Saving Mechanism: Uses FileWriter and PrintWriter to write structured text lines to disk.',
        '• Loading Mechanism: Uses FileReader and BufferedReader to restore previous records upon startup.',
        '• Pipe-Delimited Storage Format:\n   Header: SUBJECTS|Java|Mathematics|DSA\n   Row:    STUDENT|101|Tanmay Arora|85|90|75',
        '• Automatic Reload: Whenever the application is opened, loadData() parses the file and restores records into memory.'
    ];

    slide.addText(charStreamPoints.join('\n\n'), {
        x: 0.8, y: 1.6, w: 3.9, h: 3.3,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });

    // Right Column: Automated Backup (Byte Streams)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.05, w: 4.3, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.05, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('AUTOMATED BACKUP (BYTE STREAMS)', {
        x: 5.25, y: 1.13, w: 4.0, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const byteStreamPoints = [
        '• Backup File: students_backup.txt',
        '• Byte-Level Copy: Uses FileInputStream and FileOutputStream to read and write raw bytes.',
        '• Triggered by Option 8: Selecting Option 8 saves current data and immediately creates an exact backup clone.',
        '• Data Safety & Disaster Recovery: If the main students.txt file is accidentally modified or deleted, the backup file preserves the data.',
        '• Viva Distinction: Character streams handle human-readable text; Byte streams handle exact byte-by-byte duplication.'
    ];

    slide.addText(byteStreamPoints.join('\n\n'), {
        x: 5.3, y: 1.6, w: 3.9, h: 3.3,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 13: DIVISION OF WORK AMONG TEAM MEMBERS
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Division of Work Among Team Members', 'Individual Module Ownership and Responsibilities (Exact Sequence)', 13);

    const teamWork = [
        {
            num: '1',
            name: 'VISHAL',
            role: 'Student Data Management Module',
            color: NAVY,
            tasks: [
                'Created the core Student class.',
                'Implemented constructors and private getter methods.',
                'Implemented total marks and percentage calculations.',
                'Implemented academic grade calculation (A+ to F).',
                'Identified highest and lowest scoring subjects.',
                'Implemented the Reportable interface.'
            ],
            x: 0.6, y: 1.05
        },
        {
            num: '2',
            name: 'JASKARAN SINGH',
            role: 'Registration & Record Management',
            color: ROYAL,
            tasks: [
                'Implemented subject configuration module.',
                'Implemented student registration workflow.',
                'Accepted subject marks with validation.',
                'Implemented duplicate roll number prevention.',
                'Created custom InvalidMarksException (0–100 check).',
                'Implemented search and generic displayList().'
            ],
            x: 5.1, y: 1.05
        },
        {
            num: '3',
            name: 'ANSH ARORA',
            role: 'Performance Analysis & Ranking',
            color: ROYAL,
            tasks: [
                'Implemented student ranking using Java Comparator.',
                'Implemented sorting by percentage and name tie-breaker.',
                'Calculated subject-wise averages, highest, and lowest scores.',
                'Calculated overall class average percentage.',
                'Identified class topper and pass/fail counts.',
                'Generated grade distribution map using Stream API.'
            ],
            x: 0.6, y: 3.15
        },
        {
            num: '4',
            name: 'TANMAY ARORA',
            role: 'Main Program, File I/O & Integration',
            color: NAVY,
            tasks: [
                'Developed main menu and switch-case navigation.',
                'Connected all project modules into a unified tool.',
                'Implemented file persistence (FileWriter & PrintWriter).',
                'Implemented record loading (FileReader & BufferedReader).',
                'Created byte-stream backup (FileInputStream & Output).',
                'Implemented multithreaded notification & Git repo setup.'
            ],
            x: 5.1, y: 3.15
        }
    ];

    teamWork.forEach(member => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: member.x, y: member.y, w: 4.3, h: 2.0,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: member.x, y: member.y, w: 4.3, h: 0.35,
            fill: { color: member.color }
        });
        slide.addText(member.num + '. ' + member.name + ' — ' + member.role, {
            x: member.x + 0.15, y: member.y + 0.06, w: 4.0, h: 0.24,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(member.tasks.map(t => '• ' + t).join('\n'), {
            x: member.x + 0.15, y: member.y + 0.4, w: 4.0, h: 1.5,
            fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 13
        });
    });
}

// =============================================================
// SLIDE 14: LIVE PROJECT DEMONSTRATION
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Live Project Demonstration', 'Step-by-Step Terminal Execution Workflow & Sample Output Card', 14);

    // Left Column: 10 Operational Steps
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 4.4, h: 4.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 4.4, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('10-STEP DEMONSTRATION WORKFLOW', {
        x: 0.75, y: 1.12, w: 4.1, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const steps = [
        'Step 1:  Launch tool — previous records load automatically.',
        'Step 2:  Select Option 1 (Set up subjects) for the class.',
        'Step 3:  Enter subjects (e.g. Java, Mathematics, DSA).',
        'Step 4:  Select Option 2 (Add student and marks).',
        'Step 5:  Enter unique roll number and student name.',
        'Step 6:  Enter validated marks (0 to 100) for each subject.',
        'Step 7:  Console prints instant Student Performance Report.',
        'Step 8:  Select Option 5 to view overall student rankings.',
        'Step 9:  Select Option 6 & 7 for subject statistics & summary.',
        'Step 10: Select Option 8 to save records and create backup.'
    ];

    slide.addText(steps.join('\n'), {
        x: 0.8, y: 1.52, w: 4.0, h: 3.5,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 17
    });

    // Right Column: Terminal Output Preview Box
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.2, y: 1.05, w: 4.2, h: 4.1,
        fill: { color: '0A192F' }, line: { color: SKY, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.2, y: 1.05, w: 4.2, h: 0.38,
        fill: { color: '1E293B' }
    });
    slide.addText('TERMINAL REPORT PREVIEW', {
        x: 5.35, y: 1.12, w: 3.9, h: 0.25,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: SKY
    });

    const termLines = [
        '---------- STUDENT PERFORMANCE REPORT ----------',
        'Roll Number    : 101',
        'Student Name   : Tanmay Arora',
        'Java           : 88/100',
        'Mathematics    : 92/100',
        'DSA            : 85/100',
        'Total Marks    : 265/300',
        'Percentage     : 88.33%',
        'Grade          : A',
        'Highest Subject: Mathematics',
        'Lowest Subject : DSA',
        '-------------------------------------------------',
        '[Notice] Marks record added for Tanmay Arora.',
        'Data saved in students.txt',
        'Backup created as students_backup.txt'
    ];

    slide.addText(termLines.join('\n'), {
        x: 5.35, y: 1.55, w: 3.9, h: 3.4,
        fontSize: 8.5, fontFace: 'Courier New', color: 'E2E8F0', lineSpacing: 13
    });
}

// =============================================================
// SLIDE 15: CONCLUSION AND FUTURE SCOPE
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Conclusion & Future Scope', 'Project Achievements, Key Takeaways, and Future Roadmap', 15);

    // Left Column: Conclusion Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 4.3, h: 3.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 4.3, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('PROJECT CONCLUSION', {
        x: 0.75, y: 1.12, w: 4.0, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const conclusions = [
        '• Simplifies Record Management: Provides a clean, automated replacement for manual paper registers.',
        '• Eliminates Human Errors: Calculates totals, percentages, and rankings automatically.',
        '• Deep Academic Insights: Shows subject difficulty and class performance through Stream API.',
        '• Permanent Data Security: Safeguards records via file persistence and byte-stream backups.',
        '• Practical OOP Mastery: Successfully integrates Core Java, Collections, Generics, and Concurrency.'
    ];

    slide.addText(conclusions.join('\n\n'), {
        x: 0.75, y: 1.52, w: 4.0, h: 2.4,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 13
    });

    // Right Column: Future Scope Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.05, w: 4.3, h: 3.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.05, w: 4.3, h: 0.38,
        fill: { color: NAVY }
    });
    slide.addText('FUTURE SCOPE & ENHANCEMENTS', {
        x: 5.25, y: 1.12, w: 4.0, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const futureScope = [
        '• Graphical User Interface: Develop a desktop GUI using JavaFX or Swing.',
        '• Relational Database: Connect to MySQL or PostgreSQL using JDBC for large-scale data.',
        '• User Authentication: Add secure teacher and student login systems.',
        '• PDF Report Export: Generate printable, formatted student report cards in PDF.',
        '• Graphical Charts: Display visual bar charts and pie charts for class analytics.'
    ];

    slide.addText(futureScope.join('\n\n'), {
        x: 5.25, y: 1.52, w: 4.0, h: 2.4,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 13
    });

    // Bottom Banner: Thank You
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 4.18, w: 8.8, h: 1.02,
        fill: { color: NAVY }, line: { color: SKY, width: 1 }, rectRadius: 0.08
    });
    slide.addText('THANK YOU!', {
        x: 0.8, y: 4.28, w: 8.4, h: 0.45,
        fontSize: 22, fontFace: 'Arial', bold: true, color: WHITE, align: 'center'
    });
    slide.addText('Any Questions? We are happy to answer!', {
        x: 0.8, y: 4.75, w: 8.4, h: 0.35,
        fontSize: 12, fontFace: 'Arial', color: SKY, align: 'center'
    });
}

// Write the PowerPoint presentation file
const outputPath = path.join(__dirname, 'Student_Performance_Analytics_Presentation.pptx');
pptx.writeFile({ fileName: outputPath })
    .then(fileName => {
        console.log(`Presentation generated successfully: ${fileName}`);
    })
    .catch(err => {
        console.error('Error generating presentation:', err);
        process.exit(1);
    });
