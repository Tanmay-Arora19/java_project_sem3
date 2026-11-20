const fs = require('fs');
const path = require('path');
const PptxGenJS = require(process.env.APPDATA + '/npm/node_modules/pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9'; // 10 x 5.625 inches
pptx.author = 'Tanmay Arora, Vishal, Jaskaran Singh, Ansh Arora';
pptx.company = 'Chitkara University';
pptx.title = 'Student Performance Analytics Tool';

// Academic Palette: Blue, White, Black / Navy
const NAVY = '0A192F';       // Deep Navy
const ROYAL = '134074';      // Slate / Royal Blue
const ACCENT = '0077B6';     // Academic Accent Blue
const SKY = '38BDF8';        // Sky Blue Accent
const LIGHT_BG = 'F8FAFC';   // Clean Off-White Background
const CARD_BG = 'FFFFFF';    // Pure White Card Fill
const DARK_TEXT = '0F172A';  // Primary Dark Text
const BODY_TEXT = '334155';  // Secondary Text
const MUTED_TEXT = '64748B'; // Muted Grey
const BORDER_COLOR = 'CBD5E1'; // Refined Border
const GREEN = '166534';
const RED = '991B1B';
const WHITE = 'FFFFFF';

// Helper function to add slide header and footer
function addHeader(slide, titleText, subtitleText, slideNumber) {
    // Top banner
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 10, h: 0.8,
        fill: { color: NAVY }
    });
    // Accent line
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0.8, w: 10, h: 0.04,
        fill: { color: ACCENT }
    });
    // Title
    slide.addText(titleText, {
        x: 0.6, y: 0.09, w: 8.0, h: 0.4,
        fontSize: 19, fontFace: 'Arial', bold: true, color: WHITE
    });
    // Subtitle
    if (subtitleText) {
        slide.addText(subtitleText, {
            x: 0.6, y: 0.47, w: 8.0, h: 0.26,
            fontSize: 10.5, fontFace: 'Arial', color: '93C5FD'
        });
    }
    // Slide Number badge
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 8.8, y: 0.16, w: 0.65, h: 0.45,
        fill: { color: ROYAL },
        line: { color: SKY, width: 1 },
        rectRadius: 0.06
    });
    slide.addText(slideNumber.toString(), {
        x: 8.8, y: 0.16, w: 0.65, h: 0.45,
        fontSize: 13, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle'
    });

    // Bottom Footer
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 5.33, w: 10, h: 0.295,
        fill: { color: 'F1F5F9' },
        line: { color: BORDER_COLOR, width: 0.5 }
    });
    slide.addText('B.Tech CSE (AI & ML) | 3rd Semester | Chitkara University', {
        x: 0.6, y: 5.37, w: 5.0, h: 0.22,
        fontSize: 9, fontFace: 'Arial', color: MUTED_TEXT
    });
    slide.addText('Student Performance Analytics Tool', {
        x: 5.6, y: 5.37, w: 3.8, h: 0.22,
        fontSize: 9, fontFace: 'Arial', color: ROYAL, align: 'right', bold: true
    });
}

// =============================================================
// SLIDE 1: TITLE PAGE (With Faculty Guide: Rohit Sharma)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: NAVY };

    // Decorative vertical bars on left
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: 5.625, fill: { color: ACCENT } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.22, y: 0, w: 0.07, h: 5.625, fill: { color: SKY } });

    // Top Category Badge
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 0.45, w: 5.2, h: 0.38,
        fill: { color: ROYAL }, line: { color: SKY, width: 1 }, rectRadius: 0.06
    });
    slide.addText('B.TECH CSE (AI & ML) — 3RD SEMESTER COLLEGE PROJECT', {
        x: 0.8, y: 0.45, w: 5.2, h: 0.38,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: 'BAE6FD', align: 'center', valign: 'middle'
    });

    // Main Title
    slide.addText('STUDENT PERFORMANCE\nANALYTICS TOOL', {
        x: 0.8, y: 1.0, w: 8.5, h: 1.25,
        fontSize: 28, fontFace: 'Arial', bold: true, color: WHITE, lineSpacing: 35
    });

    // Subtitle
    slide.addText('A Console-Based Core Java Application for Academic Record Management & Analysis', {
        x: 0.8, y: 2.35, w: 8.5, h: 0.35,
        fontSize: 12.5, fontFace: 'Arial', color: '93C5FD'
    });

    // Divider
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.78, w: 8.4, h: 0.02, fill: { color: SKY } });

    // Left Card: Academic & Faculty Guide Info
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 2.98, w: 4.05, h: 2.25,
        fill: { color: '0F172A' }, line: { color: '334155', width: 1 }, rectRadius: 0.08
    });
    slide.addText('ACADEMIC DETAILS', {
        x: 1.0, y: 3.12, w: 3.65, h: 0.25,
        fontSize: 11.5, fontFace: 'Arial', bold: true, color: SKY
    });
    const acadLines = [
        'Subject: Java Programming',
        'Course: B.Tech CSE (AI & ML)',
        'Semester: 3rd Semester',
        'Institution: Chitkara University',
        '',
        'Faculty Guide: Rohit Sharma'
    ];
    slide.addText(acadLines.join('\n'), {
        x: 1.0, y: 3.42, w: 3.65, h: 1.7,
        fontSize: 11, fontFace: 'Arial', color: WHITE, lineSpacing: 19
    });

    // Right Card: Project Team Members (Exact Sequence)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 2.98, w: 4.05, h: 2.25,
        fill: { color: '0F172A' }, line: { color: '334155', width: 1 }, rectRadius: 0.08
    });
    slide.addText('PROJECT TEAM MEMBERS', {
        x: 5.35, y: 3.12, w: 3.65, h: 0.25,
        fontSize: 11.5, fontFace: 'Arial', bold: true, color: SKY
    });
    const teamMembers = [
        '1. Vishal           (Roll No: 2510992990)',
        '2. Jaskaran Singh   (Roll No: 2510993488)',
        '3. Ansh Arora       (Roll No: 2510992560)',
        '4. Tanmay Arora     (Roll No: 2510993577)'
    ];
    slide.addText(teamMembers.join('\n'), {
        x: 5.35, y: 3.45, w: 3.65, h: 1.65,
        fontSize: 11, fontFace: 'Arial', color: WHITE, lineSpacing: 22
    });
}

// =============================================================
// SLIDE 2: INTRODUCTION (Well-Filled & Balanced)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Introduction', 'Overview and Core Purpose of the Project', 2);

    // Top Summary Banner
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 0.98, w: 8.8, h: 0.72,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Our project is a console-based Student Performance Analytics Tool developed using Core Java. It helps teachers manage marks, calculate percentages and grades automatically, rank students, and analyze class performance with zero manual errors.', {
        x: 0.8, y: 1.05, w: 8.4, h: 0.58,
        fontSize: 11, fontFace: 'Arial', color: DARK_TEXT, lineSpacing: 16
    });

    // 4 Well-Filled Cards (2x2 Grid)
    const cards = [
        {
            title: '1. Student Record Management',
            points: [
                'Developed entirely using Core Java and OOP principles.',
                'Allows teachers to store student roll numbers and names.',
                'Supports setting up marks for multiple custom subjects.',
                'Keeps all student records organized in memory.'
            ],
            x: 0.6, y: 1.85
        },
        {
            title: '2. Automatic Calculations',
            points: [
                'Automatically calculates total marks and percentages.',
                'Assigns academic letter grades from A+ down to F.',
                'Identifies each student\'s highest and lowest subjects.',
                'Eliminates human calculation mistakes completely.'
            ],
            x: 5.1, y: 1.85
        },
        {
            title: '3. Student Ranking System',
            points: [
                'Compares overall percentages of all students in class.',
                'Ranks students automatically from highest to lowest score.',
                'Uses Java Comparator for fair, automated sorting.',
                'Resolves score ties alphabetically by student name.'
            ],
            x: 0.6, y: 3.5
        },
        {
            title: '4. Performance Analytics & Reports',
            points: [
                'Generates subject averages, top scores, and lowest marks.',
                'Provides class performance summaries and pass/fail counts.',
                'Displays grade distributions across the entire batch.',
                'Saves all student records into text files for future use.'
            ],
            x: 5.1, y: 3.5
        }
    ];

    cards.forEach(c => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: c.x, y: c.y, w: 4.3, h: 1.65,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: c.x, y: c.y, w: 4.3, h: 0.35,
            fill: { color: ROYAL }
        });
        slide.addText(c.title, {
            x: c.x + 0.15, y: c.y + 0.05, w: 4.0, h: 0.25,
            fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(c.points.map(p => '•  ' + p).join('\n'), {
            x: c.x + 0.15, y: c.y + 0.42, w: 4.0, h: 1.15,
            fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 16
        });
    });
}

// =============================================================
// SLIDE 3: PROBLEM STATEMENT (Well-Balanced & Rich Content)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Problem Statement', 'Challenges of Manual Management vs. Our Java Solution', 3);

    // Left Column: Manual Challenges
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: 'FCA5A5', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.42,
        fill: { color: RED }
    });
    slide.addText('CHALLENGES OF MANUAL MANAGEMENT', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.26,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const manualIssues = [
        { h: 'Time-Consuming Calculations', d: 'Calculating totals and percentages manually for 50+ students takes significant teacher time.' },
        { h: 'Human Calculation Errors', d: 'Manual percentage and ranking calculations often cause mistakes and grading disputes.' },
        { h: 'Difficult Subject Comparison', d: 'Comparing scores across multiple subjects on paper or spreadsheets is tedious and confusing.' },
        { h: 'Disorganized Physical Records', d: 'Paper registers can easily be lost, damaged, or misplaced over an academic semester.' },
        { h: 'Tedious Class Reports', d: 'Finding class averages, topper names, and pass/fail counts requires extra manual effort.' }
    ];

    let my = 1.55;
    manualIssues.forEach(item => {
        slide.addText('• ' + item.h, { x: 0.8, y: my, w: 3.9, h: 0.22, fontSize: 10.5, fontFace: 'Arial', bold: true, color: RED });
        slide.addText(item.d, { x: 0.95, y: my + 0.22, w: 3.75, h: 0.35, fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT });
        my += 0.64;
    });

    // Right Column: Our Java Solution
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: '86EFAC', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.42,
        fill: { color: GREEN }
    });
    slide.addText('OUR AUTOMATED JAVA SOLUTION', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.26,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const javaSolutions = [
        { h: 'Instant Automatic Calculations', d: 'Percentages, totals, and grades are calculated immediately upon mark entry with zero errors.' },
        { h: '100% Error-Free Ranking', d: 'Java Comparator sorts all students automatically based on percentage, breaking ties by name.' },
        { h: 'Fast Subject Analytics', d: 'Immediately displays subject-wise averages, top scorers, and lowest scores in clean tables.' },
        { h: 'Permanent File Storage', d: 'Records are saved in students.txt with an automatic byte-stream backup in students_backup.txt.' },
        { h: 'Simple Menu Navigation', d: 'A user-friendly numbered menu lets teachers execute all tasks in seconds with simple inputs.' }
    ];

    let sy = 1.55;
    javaSolutions.forEach(item => {
        slide.addText('✔ ' + item.h, { x: 5.3, y: sy, w: 3.9, h: 0.22, fontSize: 10.5, fontFace: 'Arial', bold: true, color: GREEN });
        slide.addText(item.d, { x: 5.45, y: sy + 0.22, w: 3.75, h: 0.35, fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT });
        sy += 0.64;
    });
}

// =============================================================
// SLIDE 4: OBJECTIVES OF THE PROJECT (Rich 8 Cards)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Project Objectives', 'Core Functional Goals and Capabilities', 4);

    const objectives = [
        { num: '01', title: 'Multi-Subject Marks Entry', desc: 'Accept and configure marks for any number of custom academic subjects dynamically.' },
        { num: '02', title: 'Automatic Percentage Calculation', desc: 'Calculate student percentages automatically using standard mathematical formulas.' },
        { num: '03', title: 'Accurate Letter Grading', desc: 'Assign standard academic letter grades (A+ to F) based on percentage thresholds.' },
        { num: '04', title: 'Subject Strengths & Weaknesses', desc: 'Identify the highest and lowest scoring subjects for every individual student.' },
        { num: '05', title: 'Automated Student Ranking', desc: 'Rank students fairly from highest to lowest percentage using Java Comparator.' },
        { num: '06', title: 'Class-Wide Performance Tracking', desc: 'Calculate overall class average, passing/failing rates, and identify the class topper.' },
        { num: '07', title: 'Subject-Wise Statistical Analysis', desc: 'Generate subject statistics including class average, top scorer, and lowest score.' },
        { num: '08', title: 'Persistent Storage & Backup', desc: 'Save all student records safely in text files and generate automated backup copies.' }
    ];

    objectives.forEach((obj, idx) => {
        const col = idx < 4 ? 0 : 1;
        const row = idx % 4;
        const x = col === 0 ? 0.6 : 5.1;
        const y = 1.05 + row * 1.03;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: y, w: 4.3, h: 0.94,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });

        // Number Badge
        slide.addShape(pptx.ShapeType.roundRect, {
            x: x + 0.12, y: y + 0.15, w: 0.52, h: 0.62,
            fill: { color: ROYAL }, rectRadius: 0.06
        });
        slide.addText(obj.num, {
            x: x + 0.12, y: y + 0.15, w: 0.52, h: 0.62,
            fontSize: 14, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle'
        });

        // Title & Description
        slide.addText(obj.title, {
            x: x + 0.75, y: y + 0.1, w: 3.45, h: 0.28,
            fontSize: 11, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(obj.desc, {
            x: x + 0.75, y: y + 0.38, w: 3.45, h: 0.5,
            fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
        });
    });
}

// =============================================================
// SLIDE 5: TECHNOLOGIES AND JAVA CONCEPTS USED (Well-Filled)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Technologies & Java Concepts Used', 'Core Java Tools and Object-Oriented Principles Implemented', 5);

    // Left Column: Tech Stack Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 3.1, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 3.1, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('DEVELOPMENT STACK', {
        x: 0.75, y: 1.1, w: 2.8, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const stackItems = [
        { label: 'Language', val: 'Java (JDK 8+)' },
        { label: 'Paradigm', val: 'Object-Oriented (OOP)' },
        { label: 'User Interface', val: 'Command Line / Terminal' },
        { label: 'Data Storage', val: 'Character Stream (students.txt)' },
        { label: 'Backup Engine', val: 'Byte Stream (students_backup.txt)' },
        { label: 'IDE Used', val: 'VS Code / IntelliJ IDEA' },
        { label: 'Version Control', val: 'Git & GitHub Repository' }
    ];

    let stkY = 1.55;
    stackItems.forEach(item => {
        slide.addText(item.label + ':', {
            x: 0.75, y: stkY, w: 2.8, h: 0.2,
            fontSize: 10, fontFace: 'Arial', bold: true, color: ACCENT
        });
        slide.addText(item.val, {
            x: 0.75, y: stkY + 0.2, w: 2.8, h: 0.28,
            fontSize: 10, fontFace: 'Arial', color: DARK_TEXT
        });
        stkY += 0.52;
    });

    // Right Column: Concepts Table
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 3.9, y: 1.02, w: 5.5, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 3.9, y: 1.02, w: 5.5, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('JAVA CONCEPT', { x: 4.05, y: 1.1, w: 2.0, h: 0.25, fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('HOW IT IS USED IN OUR PROJECT', { x: 6.1, y: 1.1, w: 3.15, h: 0.25, fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE });

    const concepts = [
        { c: 'Classes & Objects', u: 'To create and manage real-world student records.' },
        { c: 'Encapsulation', u: 'To protect student info using private fields and getters.' },
        { c: 'ArrayList', u: 'To store dynamic collections of students and subjects.' },
        { c: 'HashMap', u: 'To store marks of different subjects as key-value pairs.' },
        { c: 'Interface (Reportable)', u: 'To define a common contract for printing report cards.' },
        { c: 'Custom Exception', u: 'InvalidMarksException validates marks between 0 and 100.' },
        { c: 'Comparator', u: 'To rank students by percentage and sort ties by name.' },
        { c: 'Stream API', u: 'To calculate class averages, pass/fail counts, and grades.' },
        { c: 'Generics', u: 'To display lists of different data types using one method.' },
        { c: 'Multithreading', u: 'To run a background notification thread on record additions.' },
        { c: 'File Handling', u: 'To save records in students.txt and create backup copies.' }
    ];

    let rowY = 1.48;
    concepts.forEach((row, i) => {
        const bg = (i % 2 === 0) ? 'FFFFFF' : 'F8FAFC';
        slide.addShape(pptx.ShapeType.rect, {
            x: 3.95, y: rowY, w: 5.4, h: 0.32,
            fill: { color: bg }
        });
        slide.addText(row.c, {
            x: 4.05, y: rowY + 0.04, w: 1.95, h: 0.24,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(row.u, {
            x: 6.1, y: rowY + 0.04, w: 3.2, h: 0.24,
            fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT
        });
        rowY += 0.33;
    });
}

// =============================================================
// SLIDE 6: PROJECT ARCHITECTURE (Flowchart + Rich Explanation)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Project Architecture', 'System Workflow from User Input to File Persistence', 6);

    const steps = [
        { label: 'USER\n(Teacher / Evaluator)', color: NAVY },
        { label: 'MAIN MENU\n(Switch-Case Navigation)', color: ROYAL },
        { label: 'STUDENT DATA INPUT\n(Roll, Name, Marks)', color: ACCENT },
        { label: 'STUDENT CLASS\n(OOP Encapsulation)', color: ROYAL },
        { label: 'PERFORMANCE ENGINE\n(Total, % & Grades)', color: NAVY },
        { label: 'RANKING & ANALYTICS\n(Comparator & Streams)', color: ROYAL },
        { label: 'REPORT GENERATION\n(Reportable Interface)', color: ACCENT },
        { label: 'FILE STORAGE & BACKUP\n(students.txt & Backup)', color: NAVY }
    ];

    // Arrange in 2 rows of 4 boxes with connecting arrows
    steps.forEach((st, idx) => {
        const isRow1 = idx < 4;
        const col = isRow1 ? idx : (7 - idx);
        const x = 0.6 + col * 2.25;
        const y = isRow1 ? 1.1 : 2.65;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: y, w: 2.05, h: 1.0,
            fill: { color: st.color },
            line: { color: SKY, width: 1 },
            rectRadius: 0.08
        });
        slide.addText(st.label, {
            x: x + 0.08, y: y + 0.12, w: 1.89, h: 0.76,
            fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle'
        });

        // Connector Arrows
        if (idx < 3) {
            slide.addText('➔', { x: x + 2.05, y: y + 0.3, w: 0.25, h: 0.35, fontSize: 14, color: ACCENT, align: 'center' });
        } else if (idx === 3) {
            slide.addText('⬇', { x: x + 0.85, y: y + 1.05, w: 0.35, h: 0.45, fontSize: 16, color: ACCENT, align: 'center' });
        } else if (idx >= 4 && idx < 7) {
            slide.addText('➔', { x: x - 0.25, y: y + 0.3, w: 0.25, h: 0.35, fontSize: 14, color: ACCENT, align: 'center' });
        }
    });

    // Bottom Architecture Explanation Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 3.95, w: 8.8, h: 1.25,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 3.95, w: 8.8, h: 0.32,
        fill: { color: ROYAL }
    });
    slide.addText('HOW THE ARCHITECTURE OPERATES STEP-BY-STEP', {
        x: 0.75, y: 4.0, w: 8.5, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });
    const archExplanation = [
        '1. The user interacts through a simple Main Menu that directs requests to specific functions.',
        '2. Student inputs are validated and encapsulated safely inside the Student class, stored in ArrayList and HashMap.',
        '3. The Calculation & Analytics engines compute total marks, percentages, grades, rankings, and class-wide statistics.',
        '4. Formatted reports are displayed on screen, and all records are saved permanently to students.txt with an automatic backup.'
    ];
    slide.addText(archExplanation.join('\n'), {
        x: 0.75, y: 4.32, w: 8.5, h: 0.82,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 7: STUDENT REGISTRATION AND DATA MANAGEMENT (Well-Filled)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Student Registration & Data Management', 'How Student Information is Collected, Validated, and Stored', 7);

    const regCards = [
        {
            title: '1. Information Accepted',
            badgeCol: NAVY,
            points: [
                'Student Roll Number: Unique integer identifier for each student.',
                'Student Name: Non-empty string identifying the student.',
                'Subject Names: Defined once during subject setup (e.g. Java, Math, DSA).',
                'Subject Marks: Numerical score entered for each subject (0 to 100).'
            ],
            x: 0.6, w: 2.8
        },
        {
            title: '2. Input Validation Rules',
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
            badgeCol: ROYAL,
            points: [
                'ArrayList<Student>: Dynamically stores all student objects in memory without fixed size limits.',
                'HashMap<String, Integer>: Stores subject names as keys and marks as values for each student.',
                'ArrayList<String>: Holds the list of registered academic subjects.'
            ],
            x: 6.6, w: 2.8
        }
    ];

    regCards.forEach(c => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: c.x, y: 1.02, w: c.w, h: 3.0,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: c.x, y: 1.02, w: c.w, h: 0.38,
            fill: { color: c.badgeCol }
        });
        slide.addText(c.title, {
            x: c.x + 0.1, y: 1.1, w: c.w - 0.2, h: 0.25,
            fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(c.points.map(p => '•  ' + p).join('\n\n'), {
            x: c.x + 0.12, y: 1.48, w: c.w - 0.24, h: 2.45,
            fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
        });
    });

    // Bottom Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 4.15, w: 8.8, h: 1.05,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Key Takeaway for Viva:', {
        x: 0.8, y: 4.25, w: 3.0, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: NAVY
    });
    slide.addText('"The program strictly prevents bad data. If a user enters a roll number that already exists, it warns the user immediately. If a mark is entered outside 0 to 100, InvalidMarksException is caught and the user is prompted again without crashing the application."', {
        x: 0.8, y: 4.5, w: 8.4, h: 0.6,
        fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT, italic: true, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 8: PERCENTAGE AND GRADE CALCULATION (Well-Filled)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Percentage & Grade Calculation', 'Formulas, Step-by-Step Example, and Academic Grading System', 8);

    // Left Column: Formula & Example Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('PERCENTAGE CALCULATION & EXAMPLE', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    // Formula Box
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 1.52, w: 3.9, h: 0.75,
        fill: { color: 'F0F7FF' }, line: { color: 'BAE6FD', width: 1 }, rectRadius: 0.06
    });
    slide.addText('MATHEMATICAL FORMULA:', {
        x: 0.9, y: 1.58, w: 3.7, h: 0.2,
        fontSize: 9, fontFace: 'Arial', bold: true, color: ACCENT
    });
    slide.addText('Percentage = (Total Marks Obtained / Maximum Marks) × 100', {
        x: 0.9, y: 1.84, w: 3.7, h: 0.35,
        fontSize: 10, fontFace: 'Arial', bold: true, color: NAVY
    });

    // Practical Example
    slide.addText('PRACTICAL EXAMPLE (3 SUBJECTS):', {
        x: 0.8, y: 2.38, w: 3.9, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: DARK_TEXT
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
        x: 0.8, y: 2.65, w: 3.9, h: 2.4,
        fontSize: 9.5, fontFace: 'Courier New', color: BODY_TEXT, lineSpacing: 16
    });

    // Right Column: Grading System Table
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('OFFICIAL ACADEMIC GRADING SCALE', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    // Table Header
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.2, y: 1.52, w: 4.1, h: 0.32,
        fill: { color: 'E2E8F0' }
    });
    slide.addText('PERCENTAGE RANGE', { x: 5.3, y: 1.58, w: 1.6, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY });
    slide.addText('GRADE', { x: 7.0, y: 1.58, w: 0.8, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY });
    slide.addText('EVALUATION', { x: 7.9, y: 1.58, w: 1.3, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY });

    const grades = [
        { range: '90% – 100%', grade: 'A+', remark: 'Outstanding', bg: 'FFFFFF', color: GREEN },
        { range: '80% – 89%', grade: 'A', remark: 'Excellent', bg: 'F8FAFC', color: GREEN },
        { range: '70% – 79%', grade: 'B', remark: 'Very Good', bg: 'FFFFFF', color: NAVY },
        { range: '60% – 69%', grade: 'C', remark: 'Good', bg: 'F8FAFC', color: NAVY },
        { range: '50% – 59%', grade: 'D', remark: 'Pass', bg: 'FFFFFF', color: 'B45309' },
        { range: 'Below 50%', grade: 'F', remark: 'Fail / Reappear', bg: 'FEE2E2', color: RED }
    ];

    let gY = 1.88;
    grades.forEach(g => {
        slide.addShape(pptx.ShapeType.rect, {
            x: 5.2, y: gY, w: 4.1, h: 0.36,
            fill: { color: g.bg }
        });
        slide.addText(g.range, { x: 5.3, y: gY + 0.08, w: 1.6, h: 0.22, fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT });
        slide.addText(g.grade, { x: 7.0, y: gY + 0.08, w: 0.8, h: 0.22, fontSize: 10.5, fontFace: 'Arial', bold: true, color: g.color });
        slide.addText(g.remark, { x: 7.9, y: gY + 0.08, w: 1.3, h: 0.22, fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT });
        gY += 0.38;
    });

    slide.addText('Implementation Note:\nThe getGrade() method uses conditional if-else statements to automatically assign the letter grade immediately after computing percentage.', {
        x: 5.3, y: 4.3, w: 3.9, h: 0.75,
        fontSize: 9, fontFace: 'Arial', color: MUTED_TEXT, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 9: STUDENT RANKING USING COMPARATOR (Well-Filled)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Student Ranking Using Comparator', 'How the Program Sorts Students by Academic Performance', 9);

    // Left Column: How Comparator Works
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('HOW JAVA COMPARATOR WORKS', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const rankPoints = [
        '1. Objective Ranking: The program compares percentages of all registered students in the class.',
        '2. Descending Sort: Higher percentages receive higher ranks. The student with the highest score gets Rank 1.',
        '3. Java Comparator Interface: Implemented via Collections.sort() using a custom chained comparator.',
        '4. Alphabetical Tie-Breaker: If two students achieve identical percentages, their names are sorted alphabetically.',
        '5. Safe In-Memory Copy: Ranking is performed on a cloned ArrayList so the original registration order is preserved.'
    ];

    slide.addText(rankPoints.join('\n\n'), {
        x: 0.8, y: 1.55, w: 3.9, h: 2.5,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
    });

    // Code snippet preview
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 3.92, w: 3.9, h: 1.1,
        fill: { color: '0F172A' }, rectRadius: 0.06
    });
    slide.addText('// Java Implementation in Project:\nComparator<Student> comp = \n  Comparator.comparingDouble(Student::getPercentage)\n  .reversed().thenComparing(Student::getName);\nCollections.sort(rankedStudents, comp);', {
        x: 0.9, y: 3.98, w: 3.7, h: 0.98,
        fontSize: 8.5, fontFace: 'Courier New', color: '38BDF8', lineSpacing: 12
    });

    // Right Column: Sample Ranking Leaderboard
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('SAMPLE RANKING LEADERBOARD', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
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
            x: 5.25, y: rY, w: 4.0, h: 0.4,
            fill: { color: r.bg }, line: { color: BORDER_COLOR, width: 0.5 }, rectRadius: 0.04
        });
        slide.addText(r.rank, { x: 5.35, y: rY + 0.08, w: 0.8, h: 0.24, fontSize: 9.5, fontFace: 'Arial', bold: true, color: r.col });
        slide.addText(r.name, { x: 6.15, y: rY + 0.08, w: 1.4, h: 0.24, fontSize: 9.5, fontFace: 'Arial', bold: true, color: DARK_TEXT });
        slide.addText(r.pct + ' (' + r.grade + ')', { x: 7.55, y: rY + 0.08, w: 1.0, h: 0.24, fontSize: 9, fontFace: 'Arial', color: BODY_TEXT });
        slide.addShape(pptx.ShapeType.rect, { x: 8.55, y: rY + 0.08, w: 0.6, h: 0.24, fill: { color: r.col } });
        slide.addText(r.badge, { x: 8.55, y: rY + 0.08, w: 0.6, h: 0.24, fontSize: 7, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle' });
        rY += 0.44;
    });

    slide.addText('Key Takeaway for Viva:\n"Using Java Comparator guarantees that sorting is automated, fair, and free of manual human errors."', {
        x: 5.3, y: 4.3, w: 3.9, h: 0.7,
        fontSize: 9, fontFace: 'Arial', color: NAVY, italic: true, lineSpacing: 13
    });
}

// =============================================================
// SLIDE 10: SUBJECT-WISE STATISTICS (Well-Filled Table & Insights)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Subject-Wise Statistics', 'Analyzing Subject Difficulty, Average Performance, and Top Scorers', 10);

    // Top Summary Banner
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.0, w: 8.8, h: 0.65,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Option 6 in our program analyzes every subject across all student records. It calculates the subject average, finds the highest marks scored (with the top student name), and finds the lowest marks scored.', {
        x: 0.8, y: 1.08, w: 8.4, h: 0.48,
        fontSize: 10.5, fontFace: 'Arial', color: DARK_TEXT, lineSpacing: 15
    });

    // Center Table Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.8, w: 8.8, h: 1.9,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.8, w: 8.8, h: 0.38,
        fill: { color: NAVY }
    });

    // Table Header
    slide.addText('SUBJECT NAME', { x: 0.8, y: 1.88, w: 2.2, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('CLASS AVERAGE', { x: 3.0, y: 1.88, w: 1.8, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('HIGHEST SCORE & TOPPER', { x: 4.8, y: 1.88, w: 2.4, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('LOWEST SCORE', { x: 7.2, y: 1.88, w: 2.0, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });

    const subjectStats = [
        { sub: 'Java Programming', avg: '78.50 / 100', high: '92 / 100 (Aman Sharma)', low: '65 / 100 (Priya Verma)', bg: 'FFFFFF' },
        { sub: 'Mathematics III', avg: '81.00 / 100', high: '95 / 100 (Rahul Gupta)', low: '58 / 100 (Simran Kaur)', bg: 'F8FAFC' },
        { sub: 'Data Structures (DSA)', avg: '74.00 / 100', high: '88 / 100 (Neha Patel)', low: '60 / 100 (Rohan Mehta)', bg: 'FFFFFF' }
    ];

    let subY = 2.24;
    subjectStats.forEach(st => {
        slide.addShape(pptx.ShapeType.rect, {
            x: 0.65, y: subY, w: 8.7, h: 0.42,
            fill: { color: st.bg }
        });
        slide.addText(st.sub, { x: 0.8, y: subY + 0.1, w: 2.2, h: 0.24, fontSize: 10, fontFace: 'Arial', bold: true, color: NAVY });
        slide.addText(st.avg, { x: 3.0, y: subY + 0.1, w: 1.8, h: 0.24, fontSize: 10, fontFace: 'Arial', color: BODY_TEXT });
        slide.addText(st.high, { x: 4.8, y: subY + 0.1, w: 2.4, h: 0.24, fontSize: 10, fontFace: 'Arial', bold: true, color: GREEN });
        slide.addText(st.low, { x: 7.2, y: subY + 0.1, w: 2.0, h: 0.24, fontSize: 10, fontFace: 'Arial', color: RED });
        subY += 0.45;
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
            fontSize: 10, fontFace: 'Arial', bold: true, color: ROYAL
        });
        slide.addText(b.desc, {
            x: x + 0.15, y: 4.25, w: 2.5, h: 0.8,
            fontSize: 9, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
        });
    });
}

// =============================================================
// SLIDE 11: CLASS PERFORMANCE SUMMARY (Well-Filled)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Class Performance Summary', 'Batch-Level Academic Insights Powered by Java Stream API', 11);

    // Left Column: Key Metrics Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('OVERALL CLASS PERFORMANCE METRICS', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
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
            x: 0.8, y: mY, w: 3.9, h: 0.2,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: ACCENT
        });
        slide.addText(m.val, {
            x: 0.95, y: mY + 0.2, w: 3.75, h: 0.24,
            fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT
        });
        mY += 0.48;
    });

    // Right Column: Java Stream API Role
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('HOW JAVA STREAM API IS USED', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const streamExplanations = [
        {
            goal: '1. Class Average Calculation',
            desc: 'students.stream().mapToDouble(Student::getPercentage).average() maps each student to percentage and computes the mean.'
        },
        {
            goal: '2. Finding the Class Topper',
            desc: 'students.stream().max(Comparator.comparingDouble(...)) finds the student with highest percentage without manual loops.'
        },
        {
            goal: '3. Pass / Fail Counts',
            desc: 'students.stream().filter(s -> !s.getGrade().equals("F")).count() filters out failing students to count passed candidates.'
        },
        {
            goal: '4. Grade Distribution',
            desc: 'Collectors.groupingBy(Student::getGrade, TreeMap::new, ...) groups students by grade letter and counts frequencies.'
        }
    ];

    let stY = 1.55;
    streamExplanations.forEach(st => {
        slide.addText(st.goal, {
            x: 5.25, y: stY, w: 3.9, h: 0.22,
            fontSize: 10, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(st.desc, {
            x: 5.25, y: stY + 0.22, w: 3.9, h: 0.38,
            fontSize: 9, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 13
        });
        stY += 0.64;
    });

    slide.addText('Takeaway for Viva:\n"Stream API allows us to process large student collections cleanly with compact, readable code instead of writing multiple nested for-loops."', {
        x: 5.25, y: 4.25, w: 3.9, h: 0.75,
        fontSize: 9, fontFace: 'Arial', color: NAVY, italic: true, lineSpacing: 13
    });
}

// =============================================================
// SLIDE 12: FILE HANDLING AND BACKUP SYSTEM (Well-Filled)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'File Handling & Backup System', 'Persistent Storage Architecture and Automated Byte-Stream Backup', 12);

    // Left Column: Main Storage (Character Streams)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('DATA STORAGE (CHARACTER STREAMS)', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const charStreamPoints = [
        '• File Used: students.txt',
        '• Saving Mechanism: Uses FileWriter and PrintWriter to write structured text lines to disk.',
        '• Loading Mechanism: Uses FileReader and BufferedReader to restore previous records upon startup.',
        '• Pipe-Delimited Storage Format:\n   Header: SUBJECTS|Java|Mathematics|DSA\n   Row:    STUDENT|101|Tanmay Arora|85|90|75',
        '• Automatic Reload: Whenever the application is opened, loadData() parses the file and restores records into memory.'
    ];

    slide.addText(charStreamPoints.join('\n\n'), {
        x: 0.8, y: 1.55, w: 3.9, h: 3.4,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
    });

    // Right Column: Automated Backup (Byte Streams)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('AUTOMATED BACKUP (BYTE STREAMS)', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const byteStreamPoints = [
        '• Backup File: students_backup.txt',
        '• Byte-Level Copy: Uses FileInputStream and FileOutputStream to read and write raw bytes.',
        '• Triggered by Option 8: Selecting Option 8 saves current data and immediately creates an exact backup clone.',
        '• Data Safety & Disaster Recovery: If the main students.txt file is accidentally modified or deleted, the backup file preserves the data.',
        '• Viva Distinction: Character streams handle human-readable text; Byte streams handle exact byte-by-byte duplication.'
    ];

    slide.addText(byteStreamPoints.join('\n\n'), {
        x: 5.3, y: 1.55, w: 3.9, h: 3.4,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
    });
}

// =============================================================
// SLIDE 13: DIVISION OF WORK AMONG TEAM MEMBERS (Exact Order)
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
            x: 0.6, y: 1.02
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
            x: 5.1, y: 1.02
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
            x: member.x, y: member.y, w: 4.3, h: 2.02,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: member.x, y: member.y, w: 4.3, h: 0.35,
            fill: { color: member.color }
        });
        slide.addText(member.num + '. ' + member.name + ' — ' + member.role, {
            x: member.x + 0.15, y: member.y + 0.06, w: 4.0, h: 0.24,
            fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(member.tasks.map(t => '• ' + t).join('\n'), {
            x: member.x + 0.15, y: member.y + 0.4, w: 4.0, h: 1.55,
            fontSize: 9, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
        });
    });
}

// =============================================================
// SLIDE 14: LIVE PROJECT DEMONSTRATION (Step-by-Step + Output)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Live Project Demonstration', 'Step-by-Step Terminal Execution Workflow & Sample Output Card', 14);

    // Left Column: 10 Operational Steps
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.4, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.4, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('10-STEP DEMONSTRATION WORKFLOW', {
        x: 0.75, y: 1.1, w: 4.1, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
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
        x: 0.8, y: 1.52, w: 4.05, h: 3.5,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 18
    });

    // Right Column: Terminal Output Preview Box
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.2, y: 1.02, w: 4.2, h: 4.15,
        fill: { color: '0A192F' }, line: { color: SKY, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.2, y: 1.02, w: 4.2, h: 0.4,
        fill: { color: '1E293B' }
    });
    slide.addText('TERMINAL REPORT PREVIEW', {
        x: 5.35, y: 1.1, w: 3.9, h: 0.25,
        fontSize: 10, fontFace: 'Arial', bold: true, color: SKY
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
        x: 5.35, y: 1.52, w: 3.9, h: 3.5,
        fontSize: 9, fontFace: 'Courier New', color: 'E2E8F0', lineSpacing: 14
    });
}

// =============================================================
// SLIDE 15: CONCLUSION AND FUTURE SCOPE (Well-Filled)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Conclusion & Future Scope', 'Project Achievements, Key Takeaways, and Future Roadmap', 15);

    // Left Column: Conclusion Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 3.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('PROJECT CONCLUSION', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const conclusions = [
        '• Simplifies Record Management: Provides a clean, automated replacement for manual paper registers.',
        '• Eliminates Human Errors: Calculates totals, percentages, and rankings automatically.',
        '• Deep Academic Insights: Shows subject difficulty and class performance through Stream API.',
        '• Permanent Data Security: Safeguards records via file persistence and byte-stream backups.',
        '• Practical OOP Mastery: Successfully integrates Core Java, Collections, Generics, and Concurrency.'
    ];

    slide.addText(conclusions.join('\n\n'), {
        x: 0.75, y: 1.5, w: 4.0, h: 2.4,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });

    // Right Column: Future Scope Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 3.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.38,
        fill: { color: NAVY }
    });
    slide.addText('FUTURE SCOPE & ENHANCEMENTS', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const futureScope = [
        '• Graphical User Interface: Develop a desktop GUI using JavaFX or Swing.',
        '• Relational Database: Connect to MySQL or PostgreSQL using JDBC for large-scale data.',
        '• User Authentication: Add secure teacher and student login portals.',
        '• PDF Report Export: Generate printable, formatted student report cards in PDF.',
        '• Graphical Charts: Display visual bar charts and pie charts for class analytics.'
    ];

    slide.addText(futureScope.join('\n\n'), {
        x: 5.25, y: 1.5, w: 4.0, h: 2.4,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });

    // Bottom Banner: Thank You
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 4.15, w: 8.8, h: 1.05,
        fill: { color: NAVY }, line: { color: SKY, width: 1 }, rectRadius: 0.08
    });
    slide.addText('THANK YOU!', {
        x: 0.8, y: 4.22, w: 8.4, h: 0.45,
        fontSize: 23, fontFace: 'Arial', bold: true, color: WHITE, align: 'center'
    });
    slide.addText('Any Questions? We are happy to answer!', {
        x: 0.8, y: 4.7, w: 8.4, h: 0.35,
        fontSize: 12.5, fontFace: 'Arial', color: SKY, align: 'center'
    });
}

// Generate PPTX file
const outputPath = path.join(__dirname, 'Student_Performance_Analytics_Presentation.pptx');
pptx.writeFile({ fileName: outputPath })
    .then(fileName => {
        console.log(`Balanced presentation generated successfully: ${fileName}`);
    })
    .catch(err => {
        console.error('Error generating presentation:', err);
        process.exit(1);
    });
