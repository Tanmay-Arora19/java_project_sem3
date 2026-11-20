const fs = require('fs');
const path = require('path');
const PptxGenJS = require(process.env.APPDATA + '/npm/node_modules/pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9'; // 10 x 5.625 inches
pptx.author = 'Tanmay Arora, Vishal, Jaskaran Singh, Ansh Arora';
pptx.company = 'Chitkara University';
pptx.title = 'Student Performance Analytics Tool';

// Minimal Academic Palette
const NAVY = '0A192F';       // Deep Navy
const ROYAL = '134074';      // Slate / Royal Blue
const ACCENT = '0077B6';     // Academic Accent Blue
const SKY = '38BDF8';        // Sky Blue Accent
const LIGHT_BG = 'F8FAFC';   // Minimal Off-White
const CARD_BG = 'FFFFFF';    // Pure White Card
const DARK_TEXT = '0F172A';  // Primary Dark Text
const BODY_TEXT = '334155';  // Secondary Text
const MUTED_TEXT = '64748B'; // Muted Grey
const BORDER_COLOR = 'E2E8F0'; // Soft Border
const GREEN = '166534';
const RED = '991B1B';
const WHITE = 'FFFFFF';

// Helper function to add minimal slide header and footer
function addHeader(slide, titleText, subtitleText, slideNumber) {
    // Top banner
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 10, h: 0.75,
        fill: { color: NAVY }
    });
    // Accent line
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0.75, w: 10, h: 0.04,
        fill: { color: ACCENT }
    });
    // Title
    slide.addText(titleText, {
        x: 0.7, y: 0.08, w: 7.8, h: 0.38,
        fontSize: 18, fontFace: 'Arial', bold: true, color: WHITE
    });
    // Subtitle
    if (subtitleText) {
        slide.addText(subtitleText, {
            x: 0.7, y: 0.44, w: 7.8, h: 0.24,
            fontSize: 10, fontFace: 'Arial', color: '93C5FD'
        });
    }
    // Slide Number badge
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 8.8, y: 0.16, w: 0.65, h: 0.42,
        fill: { color: ROYAL },
        line: { color: SKY, width: 1 },
        rectRadius: 0.06
    });
    slide.addText(slideNumber.toString(), {
        x: 8.8, y: 0.16, w: 0.65, h: 0.42,
        fontSize: 12, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle'
    });

    // Bottom Footer
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 5.35, w: 10, h: 0.275,
        fill: { color: 'F1F5F9' },
        line: { color: BORDER_COLOR, width: 0.5 }
    });
    slide.addText('B.Tech CSE (AI & ML) | 3rd Semester | Chitkara University', {
        x: 0.7, y: 5.38, w: 5.0, h: 0.22,
        fontSize: 8.5, fontFace: 'Arial', color: MUTED_TEXT
    });
    slide.addText('Student Performance Analytics Tool', {
        x: 5.7, y: 5.38, w: 3.6, h: 0.22,
        fontSize: 8.5, fontFace: 'Arial', color: ROYAL, align: 'right', bold: true
    });
}

// =============================================================
// SLIDE 1: TITLE PAGE (With Faculty Guide: Rohit Sharma)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: NAVY };

    // Decorative vertical bar on left
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.2, h: 5.625, fill: { color: ACCENT } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.2, y: 0, w: 0.06, h: 5.625, fill: { color: SKY } });

    // Top Category Badge
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 0.5, w: 4.8, h: 0.35,
        fill: { color: ROYAL }, line: { color: SKY, width: 1 }, rectRadius: 0.06
    });
    slide.addText('B.TECH CSE (AI & ML) — 3RD SEMESTER PROJECT', {
        x: 0.8, y: 0.5, w: 4.8, h: 0.35,
        fontSize: 9, fontFace: 'Arial', bold: true, color: 'BAE6FD', align: 'center', valign: 'middle'
    });

    // Main Title
    slide.addText('STUDENT PERFORMANCE\nANALYTICS TOOL', {
        x: 0.8, y: 1.05, w: 8.5, h: 1.25,
        fontSize: 27, fontFace: 'Arial', bold: true, color: WHITE, lineSpacing: 34
    });

    // Subtitle
    slide.addText('A Console-Based Core Java Application for Academic Record Management & Analysis', {
        x: 0.8, y: 2.38, w: 8.5, h: 0.35,
        fontSize: 12, fontFace: 'Arial', color: '93C5FD'
    });

    // Minimal Divider
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.8, w: 8.4, h: 0.02, fill: { color: SKY } });

    // Left Card: Academic & Faculty Guide Info
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 3.0, w: 4.0, h: 2.2,
        fill: { color: '0F172A' }, line: { color: '334155', width: 1 }, rectRadius: 0.08
    });
    slide.addText('ACADEMIC DETAILS', {
        x: 1.0, y: 3.15, w: 3.6, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: SKY
    });
    
    // Academic lines with Faculty Guide clearly highlighted
    const acadLines = [
        'Subject: Java Programming',
        'Course: B.Tech CSE (AI & ML)',
        'Institution: Chitkara University',
        '',
        'Faculty Guide: Rohit Sharma'
    ];
    slide.addText(acadLines.join('\n'), {
        x: 1.0, y: 3.48, w: 3.6, h: 1.6,
        fontSize: 10.5, fontFace: 'Arial', color: WHITE, lineSpacing: 20
    });

    // Right Card: Project Team Members (Exact Order)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.2, y: 3.0, w: 4.0, h: 2.2,
        fill: { color: '0F172A' }, line: { color: '334155', width: 1 }, rectRadius: 0.08
    });
    slide.addText('PROJECT TEAM MEMBERS', {
        x: 5.4, y: 3.15, w: 3.6, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: SKY
    });
    const teamMembers = [
        '1. Vishal           (Roll No: 2510992990)',
        '2. Jaskaran Singh   (Roll No: 2510993488)',
        '3. Ansh Arora       (Roll No: 2510992560)',
        '4. Tanmay Arora     (Roll No: 2510993577)'
    ];
    slide.addText(teamMembers.join('\n'), {
        x: 5.4, y: 3.48, w: 3.6, h: 1.6,
        fontSize: 10.5, fontFace: 'Arial', color: WHITE, lineSpacing: 22
    });
}

// =============================================================
// SLIDE 2: INTRODUCTION (Minimal & Clean)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Introduction', 'Overview of the Student Performance Analytics Tool', 2);

    // Top Simple Banner
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.0, w: 8.6, h: 0.65,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('A terminal-based Core Java application that helps teachers store marks, compute grades, rank students, and generate academic reports automatically.', {
        x: 0.9, y: 1.1, w: 8.2, h: 0.45,
        fontSize: 11.5, fontFace: 'Arial', color: DARK_TEXT, align: 'center', valign: 'middle'
    });

    // 4 Minimal Cards (2x2 Grid)
    const cards = [
        {
            title: '1. Record Management',
            points: [
                'Built using Core Java.',
                'Stores roll numbers and student names.',
                'Supports custom academic subjects.'
            ],
            x: 0.7, y: 1.85
        },
        {
            title: '2. Automatic Calculations',
            points: [
                'Computes total marks and percentages.',
                'Assigns grades from A+ down to F.',
                'Identifies highest & lowest subjects.'
            ],
            x: 5.15, y: 1.85
        },
        {
            title: '3. Student Ranking',
            points: [
                'Compares overall percentages.',
                'Ranks students automatically.',
                'Uses Java Comparator for fair sorting.'
            ],
            x: 0.7, y: 3.45
        },
        {
            title: '4. Performance Reports',
            points: [
                'Generates subject averages and toppers.',
                'Provides overall class summary.',
                'Saves data permanently in text files.'
            ],
            x: 5.15, y: 3.45
        }
    ];

    cards.forEach(c => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: c.x, y: c.y, w: 4.15, h: 1.4,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: c.x, y: c.y, w: 4.15, h: 0.32,
            fill: { color: ROYAL }
        });
        slide.addText(c.title, {
            x: c.x + 0.15, y: c.y + 0.04, w: 3.85, h: 0.25,
            fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(c.points.map(p => '•  ' + p).join('\n'), {
            x: c.x + 0.15, y: c.y + 0.38, w: 3.85, h: 0.95,
            fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 18
        });
    });
}

// =============================================================
// SLIDE 3: PROBLEM STATEMENT (Minimal Comparison)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Problem Statement', 'Manual Challenges vs. Our Automated Solution', 3);

    // Left Column: Manual Challenges
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: 'FCA5A5', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.05, w: 4.15, h: 0.4,
        fill: { color: RED }
    });
    slide.addText('MANUAL CHALLENGES', {
        x: 0.85, y: 1.13, w: 3.8, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const manualIssues = [
        'Calculating percentages manually takes too much time.',
        'Manual ranking often causes human calculation errors.',
        'Comparing marks across multiple subjects is difficult.',
        'Paper records can easily be lost or damaged.',
        'Preparing class performance reports requires extra effort.'
    ];

    slide.addText(manualIssues.map(p => '•  ' + p).join('\n\n'), {
        x: 0.9, y: 1.65, w: 3.75, h: 3.2,
        fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 17
    });

    // Right Column: Our Java Solution
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: '86EFAC', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.15, y: 1.05, w: 4.15, h: 0.4,
        fill: { color: GREEN }
    });
    slide.addText('OUR JAVA SOLUTION', {
        x: 5.3, y: 1.13, w: 3.8, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const javaSolutions = [
        'Percentages and grades are calculated instantly.',
        'Java Comparator ensures 100% accurate, fair ranking.',
        'Displays subject averages and toppers in clean tables.',
        'Permanent file storage in students.txt with backup.',
        'Simple numbered menu connects all functions smoothly.'
    ];

    slide.addText(javaSolutions.map(s => '✔  ' + s).join('\n\n'), {
        x: 5.35, y: 1.65, w: 3.75, h: 3.2,
        fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 17
    });
}

// =============================================================
// SLIDE 4: OBJECTIVES OF THE PROJECT (Minimal Chips)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Project Objectives', 'Core Goals of the Application', 4);

    const objectives = [
        '1. Accept marks for multiple subjects dynamically.',
        '2. Calculate student percentages automatically.',
        '3. Assign standard letter grades (A+ to F).',
        '4. Identify highest and lowest scoring subjects.',
        '5. Rank students fairly using Java Comparator.',
        '6. Calculate overall class averages and find the topper.',
        '7. Generate subject-wise statistics and top scorers.',
        '8. Save student records and create automated backups.'
    ];

    objectives.forEach((obj, idx) => {
        const col = idx < 4 ? 0 : 1;
        const row = idx % 4;
        const x = col === 0 ? 0.7 : 5.15;
        const y = 1.05 + row * 1.0;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: y, w: 4.15, h: 0.85,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });

        // Left accent stripe
        slide.addShape(pptx.ShapeType.rect, {
            x: x, y: y, w: 0.1, h: 0.85,
            fill: { color: ROYAL }
        });

        slide.addText(obj, {
            x: x + 0.22, y: y + 0.12, w: 3.8, h: 0.6,
            fontSize: 10.5, fontFace: 'Arial', bold: true, color: NAVY, valign: 'middle'
        });
    });
}

// =============================================================
// SLIDE 5: TECHNOLOGIES AND JAVA CONCEPTS (Minimal Table)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Technologies & Java Concepts Used', 'Tools and Core Java Principles Implemented', 5);

    // Left Column: Tech Stack (Minimal Card)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.05, w: 2.8, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.05, w: 2.8, h: 0.35,
        fill: { color: NAVY }
    });
    slide.addText('DEVELOPMENT STACK', {
        x: 0.85, y: 1.12, w: 2.5, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const stack = [
        'Language: Java (JDK 8+)',
        'Paradigm: Object-Oriented',
        'Interface: Command Line',
        'Storage: Text File (students.txt)',
        'Backup: Byte Stream Copy',
        'IDE: VS Code / IntelliJ IDEA'
    ];
    slide.addText(stack.map(s => '•  ' + s).join('\n\n'), {
        x: 0.85, y: 1.55, w: 2.5, h: 3.3,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 18
    });

    // Right Column: Java Concepts Table
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 3.7, y: 1.05, w: 5.6, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 3.7, y: 1.05, w: 5.6, h: 0.35,
        fill: { color: ROYAL }
    });
    slide.addText('JAVA CONCEPT', { x: 3.85, y: 1.12, w: 2.0, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('HOW IT IS USED IN OUR PROJECT', { x: 5.85, y: 1.12, w: 3.3, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });

    const concepts = [
        { c: 'Classes & Objects', u: 'To create and manage student records.' },
        { c: 'Encapsulation', u: 'To protect student info using private variables.' },
        { c: 'ArrayList', u: 'To store dynamic lists of students and subjects.' },
        { c: 'HashMap', u: 'To store subject names and marks as key-values.' },
        { c: 'Interface (Reportable)', u: 'To define student report card printing.' },
        { c: 'Custom Exception', u: 'To validate marks between 0 and 100.' },
        { c: 'Comparator', u: 'To rank students according to percentages.' },
        { c: 'Stream API', u: 'To calculate class performance statistics.' },
        { c: 'Generics', u: 'To display lists of different data types.' },
        { c: 'Multithreading', u: 'To display record-added notifications.' },
        { c: 'File Handling', u: 'To save and retrieve student records.' }
    ];

    let rowY = 1.45;
    concepts.forEach((row, i) => {
        const bg = (i % 2 === 0) ? 'FFFFFF' : 'F8FAFC';
        slide.addShape(pptx.ShapeType.rect, {
            x: 3.75, y: rowY, w: 5.5, h: 0.31,
            fill: { color: bg }
        });
        slide.addText(row.c, {
            x: 3.85, y: rowY + 0.04, w: 1.95, h: 0.23,
            fontSize: 8.5, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(row.u, {
            x: 5.85, y: rowY + 0.04, w: 3.35, h: 0.23,
            fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT
        });
        rowY += 0.32;
    });
}

// =============================================================
// SLIDE 6: PROJECT ARCHITECTURE (Clean Minimal Flowchart)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Project Architecture', 'System Workflow from User Input to File Storage', 6);

    const steps = [
        { label: 'USER', sub: 'Teacher' },
        { label: 'MAIN MENU', sub: 'Options 1-9' },
        { label: 'STUDENT INPUT', sub: 'Roll, Name, Marks' },
        { label: 'STUDENT CLASS', sub: 'Encapsulated Model' },
        { label: 'CALCULATIONS', sub: 'Total, % & Grade' },
        { label: 'RANKING & STATS', sub: 'Comparator & Streams' },
        { label: 'REPORTS', sub: 'Reportable Interface' },
        { label: 'FILE STORAGE', sub: 'students.txt & Backup' }
    ];

    // 2 Rows of 4 blocks
    steps.forEach((st, idx) => {
        const isRow1 = idx < 4;
        const col = isRow1 ? idx : (7 - idx);
        const x = 0.7 + col * 2.2;
        const y = isRow1 ? 1.2 : 2.7;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: y, w: 2.0, h: 0.95,
            fill: { color: isRow1 ? NAVY : ROYAL },
            line: { color: SKY, width: 1 },
            rectRadius: 0.06
        });
        slide.addText(st.label, {
            x: x, y: y + 0.16, w: 2.0, h: 0.35,
            fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE, align: 'center'
        });
        slide.addText(st.sub, {
            x: x, y: y + 0.48, w: 2.0, h: 0.3,
            fontSize: 8.5, fontFace: 'Arial', color: '93C5FD', align: 'center'
        });

        // Arrows
        if (idx < 3) {
            slide.addText('➔', { x: x + 2.0, y: y + 0.28, w: 0.2, h: 0.35, fontSize: 13, color: ACCENT, align: 'center' });
        } else if (idx === 3) {
            slide.addText('⬇', { x: x + 0.85, y: y + 1.0, w: 0.3, h: 0.45, fontSize: 16, color: ACCENT, align: 'center' });
        } else if (idx >= 4 && idx < 7) {
            slide.addText('➔', { x: x - 0.2, y: y + 0.28, w: 0.2, h: 0.35, fontSize: 13, color: ACCENT, align: 'center' });
        }
    });

    // Minimal Explanation Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 4.0, w: 8.6, h: 1.05,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('How the System Operates:', {
        x: 0.9, y: 4.12, w: 8.2, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: NAVY
    });
    slide.addText('The Main Menu connects all project functions. Student inputs are encapsulated in the Student class, calculations and rankings are performed automatically, formatted reports are printed, and records are saved to files.', {
        x: 0.9, y: 4.38, w: 8.2, h: 0.55,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
    });
}

// =============================================================
// SLIDE 7: STUDENT REGISTRATION AND DATA MANAGEMENT (Minimal)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Student Registration & Data Management', 'Accepting, Validating, and Storing Records', 7);

    const cards = [
        {
            title: '1. Information Accepted',
            color: NAVY,
            points: [
                'Student roll number.',
                'Student name.',
                'Multiple subject names.',
                'Marks for every subject.'
            ],
            x: 0.7
        },
        {
            title: '2. Input Validation',
            color: GREEN,
            points: [
                'Validation of marks between 0 and 100.',
                'Uses InvalidMarksException.',
                'Prevention of duplicate roll numbers.',
                'Prompts user again on error.'
            ],
            x: 3.65
        },
        {
            title: '3. Data Structures Used',
            color: ROYAL,
            points: [
                'ArrayList stores student objects.',
                'HashMap stores subject-wise marks.',
                'ArrayList stores subject names.',
                'Organized in-memory storage.'
            ],
            x: 6.6
        }
    ];

    cards.forEach(c => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: c.x, y: 1.1, w: 2.7, h: 2.8,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: c.x, y: 1.1, w: 2.7, h: 0.35,
            fill: { color: c.color }
        });
        slide.addText(c.title, {
            x: c.x + 0.1, y: 1.17, w: 2.5, h: 0.22,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(c.points.map(p => '•  ' + p).join('\n\n'), {
            x: c.x + 0.12, y: 1.6, w: 2.45, h: 2.1,
            fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 17
        });
    });

    // Bottom Summary
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 4.1, w: 8.6, h: 0.95,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Key Takeaway:', {
        x: 0.9, y: 4.22, w: 8.2, h: 0.2,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: NAVY
    });
    slide.addText('ArrayList stores the collection of Student objects, while a HashMap inside each Student maps subject names to their marks (0 to 100).', {
        x: 0.9, y: 4.45, w: 8.2, h: 0.48,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT
    });
}

// =============================================================
// SLIDE 8: PERCENTAGE AND GRADE CALCULATION (Minimal Layout)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Percentage & Grade Calculation', 'Formulas, Example, and Grading System', 8);

    // Left Column: Formula & Example
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('PERCENTAGE CALCULATION', {
        x: 0.85, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    slide.addText('Formula:\nPercentage = (Total Marks / Maximum Marks) × 100', {
        x: 0.85, y: 1.55, w: 3.8, h: 0.55,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: NAVY, lineSpacing: 15
    });

    const example = [
        'Example:',
        '• Java = 85',
        '• Mathematics = 90',
        '• DSA = 75',
        '---------------------------------------',
        'Total Marks = 250',
        'Maximum Marks = 300',
        'Percentage = (250 / 300) × 100 = 83.33%',
        'Grade = A'
    ];
    slide.addText(example.join('\n'), {
        x: 0.85, y: 2.2, w: 3.8, h: 2.6,
        fontSize: 9.5, fontFace: 'Courier New', color: BODY_TEXT, lineSpacing: 16
    });

    // Right Column: Grading System Table
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.15, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: NAVY }
    });
    slide.addText('GRADING SYSTEM', {
        x: 5.3, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const gradeRows = [
        { r: '90 – 100%', g: 'A+', rem: 'Outstanding', col: GREEN },
        { r: '80 – 89%',  g: 'A',  rem: 'Excellent', col: GREEN },
        { r: '70 – 79%',  g: 'B',  rem: 'Very Good', col: NAVY },
        { r: '60 – 69%',  g: 'C',  rem: 'Good', col: NAVY },
        { r: '50 – 59%',  g: 'D',  rem: 'Pass', col: 'B45309' },
        { r: 'Below 50%', g: 'F',  rem: 'Fail', col: RED }
    ];

    let gy = 1.55;
    gradeRows.forEach(gr => {
        slide.addText(gr.r, { x: 5.3, y: gy, w: 1.5, h: 0.24, fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT });
        slide.addText(gr.g, { x: 6.9, y: gy, w: 0.8, h: 0.24, fontSize: 10.5, fontFace: 'Arial', bold: true, color: gr.col });
        slide.addText(gr.rem, { x: 7.7, y: gy, w: 1.4, h: 0.24, fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT });
        gy += 0.38;
    });

    slide.addText('The program automatically calculates grades using conditional if-else statements.', {
        x: 5.3, y: 4.25, w: 3.8, h: 0.6,
        fontSize: 9, fontFace: 'Arial', color: MUTED_TEXT, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 9: STUDENT RANKING USING COMPARATOR (Minimal)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Student Ranking Using Comparator', 'Sorting Students by Academic Performance', 9);

    // Left Column: How it Works
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: NAVY }
    });
    slide.addText('HOW RANKING WORKS', {
        x: 0.85, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const rankPoints = [
        'The program compares percentages of all students.',
        'Students with higher percentages receive higher ranks.',
        'Java Comparator is used for sorting in descending order.',
        'If two students have equal percentages, names are used for alphabetical ordering.',
        'The highest percentage receives Rank 1.'
    ];

    slide.addText(rankPoints.map(p => '•  ' + p).join('\n\n'), {
        x: 0.85, y: 1.6, w: 3.8, h: 3.2,
        fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 18
    });

    // Right Column: Small Example
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.15, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('EXAMPLE RANKING', {
        x: 5.3, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const sampleRanks = [
        { rank: '1', name: 'Student A', pct: '92%', grade: 'A+' },
        { rank: '2', name: 'Student B', pct: '87%', grade: 'A' },
        { rank: '3', name: 'Student C', pct: '83%', grade: 'A' },
        { rank: '4', name: 'Student D', pct: '74%', grade: 'B' }
    ];

    let ry = 1.65;
    sampleRanks.forEach(sr => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.3, y: ry, w: 3.85, h: 0.45,
            fill: { color: sr.rank === '1' ? 'DCFCE7' : 'F8FAFC' },
            line: { color: BORDER_COLOR, width: 0.5 },
            rectRadius: 0.04
        });
        slide.addText('Rank ' + sr.rank + ': ' + sr.name + '  —  ' + sr.pct + ' (Grade ' + sr.grade + ')', {
            x: 5.45, y: ry + 0.08, w: 3.5, h: 0.28,
            fontSize: 10, fontFace: 'Arial', bold: (sr.rank === '1'), color: DARK_TEXT
        });
        ry += 0.55;
    });

    slide.addText('Comparator eliminates manual errors and saves hours of teacher calculation.', {
        x: 5.3, y: 4.1, w: 3.8, h: 0.6,
        fontSize: 9, fontFace: 'Arial', color: NAVY, italic: true
    });
}

// =============================================================
// SLIDE 10: SUBJECT-WISE STATISTICS (Minimal Table)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Subject-Wise Statistics', 'Analyzing Subject Difficulty and Student Performance', 10);

    // Top Summary
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.0, w: 8.6, h: 0.55,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.06
    });
    slide.addText('The program analyses every subject to find the average marks, highest marks with student name, and lowest marks with student name.', {
        x: 0.85, y: 1.06, w: 8.3, h: 0.42,
        fontSize: 10, fontFace: 'Arial', color: DARK_TEXT, align: 'center', valign: 'middle'
    });

    // Center Clean Table
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.7, w: 8.6, h: 2.1,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.7, w: 8.6, h: 0.38,
        fill: { color: NAVY }
    });

    slide.addText('SUBJECT', { x: 0.9, y: 1.78, w: 2.2, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('AVERAGE', { x: 3.1, y: 1.78, w: 1.6, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('HIGHEST MARKS', { x: 4.8, y: 1.78, w: 2.3, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('LOWEST MARKS', { x: 7.2, y: 1.78, w: 1.9, h: 0.22, fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE });

    const subStats = [
        { s: 'Java', avg: '78.50', h: '92 (Aman)', l: '65 (Priya)' },
        { s: 'Mathematics', avg: '81.00', h: '95 (Rahul)', l: '58 (Simran)' },
        { s: 'DSA', avg: '74.00', h: '88 (Neha)', l: '60 (Rohan)' }
    ];

    let sty = 2.18;
    subStats.forEach(st => {
        slide.addText(st.s, { x: 0.9, y: sty, w: 2.2, h: 0.25, fontSize: 10, fontFace: 'Arial', bold: true, color: NAVY });
        slide.addText(st.avg, { x: 3.1, y: sty, w: 1.6, h: 0.25, fontSize: 10, fontFace: 'Arial', color: BODY_TEXT });
        slide.addText(st.h, { x: 4.8, y: sty, w: 2.3, h: 0.25, fontSize: 10, fontFace: 'Arial', bold: true, color: GREEN });
        slide.addText(st.l, { x: 7.2, y: sty, w: 1.9, h: 0.25, fontSize: 10, fontFace: 'Arial', color: RED });
        sty += 0.45;
    });

    // Bottom Insight
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 4.0, w: 8.6, h: 1.05,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Why this helps teachers:', {
        x: 0.9, y: 4.12, w: 8.2, h: 0.22,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: NAVY
    });
    slide.addText('This helps teachers quickly understand which subjects students perform well in and which subjects they find difficult so that extra guidance can be provided.', {
        x: 0.9, y: 4.38, w: 8.2, h: 0.55,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 11: CLASS PERFORMANCE SUMMARY (Minimal)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Class Performance Summary', 'Overall Class Analytics using Java Stream API', 11);

    // Left Column: Key Metrics
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: NAVY }
    });
    slide.addText('CLASS METRICS INCLUDED', {
        x: 0.85, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const classPoints = [
        'Total number of students in the class.',
        'Overall class average percentage.',
        'Number of passed students.',
        'Number of failed students.',
        'Name of the class topper.',
        'Topper\'s percentage.',
        'Grade distribution (A+, A, B, C, D, F).'
    ];

    slide.addText(classPoints.map(p => '•  ' + p).join('\n\n'), {
        x: 0.85, y: 1.55, w: 3.8, h: 3.3,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 17
    });

    // Right Column: Java Stream API Role
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.15, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('HOW JAVA STREAM API IS USED', {
        x: 5.3, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const streamPoints = [
        'average(): To calculate the class average percentage.',
        'max(): To find the student with the highest marks.',
        'filter(): To count passed and failed students.',
        'Collectors.groupingBy(): To count how many students received each grade.'
    ];

    slide.addText(streamPoints.map(p => '✔  ' + p).join('\n\n'), {
        x: 5.3, y: 1.6, w: 3.8, h: 2.4,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 17
    });

    slide.addText('Java Stream API allows fast calculations in very few lines of code without complex loops.', {
        x: 5.3, y: 4.1, w: 3.8, h: 0.6,
        fontSize: 9, fontFace: 'Arial', color: NAVY, italic: true
    });
}

// =============================================================
// SLIDE 12: FILE HANDLING AND BACKUP SYSTEM (Minimal)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'File Handling & Backup System', 'How Student Records are Saved and Protected', 12);

    // Left Column: Main Storage
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('DATA STORAGE (students.txt)', {
        x: 0.85, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const filePoints = [
        'Students\' information is stored in students.txt.',
        'FileWriter and PrintWriter are used to save student records.',
        'BufferedReader and FileReader are used to load previous records.',
        'Data remains available when the application is restarted.'
    ];

    slide.addText(filePoints.map(p => '•  ' + p).join('\n\n'), {
        x: 0.85, y: 1.65, w: 3.8, h: 3.0,
        fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 18
    });

    // Right Column: Backup System
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.15, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: NAVY }
    });
    slide.addText('BACKUP SYSTEM (students_backup.txt)', {
        x: 5.3, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const backupPoints = [
        'Backup is stored in students_backup.txt.',
        'FileInputStream and FileOutputStream are used to create the backup.',
        'Copies raw bytes to ensure an exact duplicate.',
        'Prevents data loss if the main file is ever damaged.'
    ];

    slide.addText(backupPoints.map(p => '✔  ' + p).join('\n\n'), {
        x: 5.3, y: 1.65, w: 3.8, h: 3.0,
        fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 18
    });
}

// =============================================================
// SLIDE 13: DIVISION OF WORK AMONG TEAM MEMBERS (Exact Order)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Division of Work Among Team Members', 'Individual Contribution of All Four Members (Exact Sequence)', 13);

    const team = [
        {
            num: '1',
            name: 'VISHAL',
            role: 'Student Data Management',
            color: NAVY,
            tasks: [
                'Created the Student class.',
                'Implemented constructors and getter methods.',
                'Implemented total marks and percentage calculations.',
                'Implemented grade calculation.',
                'Identified highest and lowest scoring subjects.',
                'Implemented the Reportable interface.'
            ],
            x: 0.7, y: 1.05
        },
        {
            num: '2',
            name: 'JASKARAN SINGH',
            role: 'Student Registration & Record Management',
            color: ROYAL,
            tasks: [
                'Implemented subject setup.',
                'Implemented student registration.',
                'Accepted student marks.',
                'Added input validation.',
                'Prevented duplicate roll numbers.',
                'Implemented searching and custom exception.'
            ],
            x: 5.15, y: 1.05
        },
        {
            num: '3',
            name: 'ANSH ARORA',
            role: 'Performance Analysis and Ranking',
            color: ROYAL,
            tasks: [
                'Implemented student ranking using Comparator.',
                'Implemented sorting.',
                'Calculated subject-wise averages.',
                'Identified highest and lowest marks.',
                'Calculated class averages and topper.',
                'Generated grade distribution.'
            ],
            x: 0.7, y: 3.15
        },
        {
            num: '4',
            name: 'TANMAY ARORA',
            role: 'Main Program, File Handling & Integration',
            color: NAVY,
            tasks: [
                'Developed the main menu.',
                'Connected all project functions.',
                'Implemented switch-case navigation.',
                'Implemented file handling and saving.',
                'Created the backup system.',
                'Implemented notification thread and testing.'
            ],
            x: 5.15, y: 3.15
        }
    ];

    team.forEach(m => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: m.x, y: m.y, w: 4.15, h: 1.95,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: m.x, y: m.y, w: 4.15, h: 0.32,
            fill: { color: m.color }
        });
        slide.addText(m.num + '. ' + m.name + ' — ' + m.role, {
            x: m.x + 0.15, y: m.y + 0.05, w: 3.85, h: 0.22,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(m.tasks.map(t => '•  ' + t).join('\n'), {
            x: m.x + 0.15, y: m.y + 0.38, w: 3.85, h: 1.5,
            fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 13
        });
    });
}

// =============================================================
// SLIDE 14: LIVE PROJECT DEMONSTRATION (Minimal)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Live Project Demonstration', 'Execution Steps and Sample Terminal Output', 14);

    // Left Column: Steps
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('EXECUTION SEQUENCE (10 STEPS)', {
        x: 0.85, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const steps = [
        'STEP 1: Open the application.',
        'STEP 2: Select the subject setup option.',
        'STEP 3: Enter subject names (Java, Math, DSA).',
        'STEP 4: Select Add Student.',
        'STEP 5: Enter student roll number and name.',
        'STEP 6: Enter marks for all subjects.',
        'STEP 7: Display the generated student report.',
        'STEP 8: Use ranking option to compare students.',
        'STEP 9: Display subject statistics & class summary.',
        'STEP 10: Save records and create a backup.'
    ];

    slide.addText(steps.join('\n'), {
        x: 0.85, y: 1.55, w: 3.8, h: 3.3,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 16
    });

    // Right Column: Minimal Terminal Output
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 1.05, w: 4.15, h: 4.0,
        fill: { color: '0A192F' }, line: { color: SKY, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.15, y: 1.05, w: 4.15, h: 0.38,
        fill: { color: '1E293B' }
    });
    slide.addText('SAMPLE TERMINAL OUTPUT', {
        x: 5.3, y: 1.13, w: 3.8, h: 0.22,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: SKY
    });

    const termText = [
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

    slide.addText(termText.join('\n'), {
        x: 5.3, y: 1.55, w: 3.85, h: 3.3,
        fontSize: 8.5, fontFace: 'Courier New', color: 'E2E8F0', lineSpacing: 13
    });
}

// =============================================================
// SLIDE 15: CONCLUSION AND FUTURE SCOPE (Minimal)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Conclusion & Future Scope', 'Summary and Upcoming Improvements', 15);

    // Left Column: Conclusion
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 1.05, w: 4.15, h: 2.9,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.7, y: 1.05, w: 4.15, h: 0.35,
        fill: { color: ROYAL }
    });
    slide.addText('CONCLUSION', {
        x: 0.85, y: 1.12, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const conclusion = [
        'Our project makes student record management easier.',
        'It reduces manual calculations.',
        'It generates percentages and grades automatically.',
        'It helps teachers compare student performance.',
        'It provides class and subject-wise statistics.',
        'It stores records for future use.'
    ];

    slide.addText(conclusion.map(c => '•  ' + c).join('\n'), {
        x: 0.85, y: 1.5, w: 3.8, h: 2.3,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
    });

    // Right Column: Future Scope
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 1.05, w: 4.15, h: 2.9,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.15, y: 1.05, w: 4.15, h: 0.35,
        fill: { color: NAVY }
    });
    slide.addText('FUTURE SCOPE', {
        x: 5.3, y: 1.12, w: 3.8, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
    });

    const future = [
        'Adding a graphical user interface (GUI).',
        'Connecting the application to a database.',
        'Adding student login functionality.',
        'Generating PDF performance reports.',
        'Adding graphical performance charts.'
    ];

    slide.addText(future.map(f => '✔  ' + f).join('\n'), {
        x: 5.3, y: 1.5, w: 3.8, h: 2.3,
        fontSize: 8.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 17
    });

    // Bottom Banner: Thank You!
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.7, y: 4.1, w: 8.6, h: 1.05,
        fill: { color: NAVY }, line: { color: SKY, width: 1 }, rectRadius: 0.08
    });
    slide.addText('THANK YOU!', {
        x: 0.9, y: 4.18, w: 8.2, h: 0.45,
        fontSize: 22, fontFace: 'Arial', bold: true, color: WHITE, align: 'center'
    });
    slide.addText('Any Questions?', {
        x: 0.9, y: 4.65, w: 8.2, h: 0.35,
        fontSize: 12, fontFace: 'Arial', color: SKY, align: 'center'
    });
}

// Generate PPTX file
const outputPath = path.join(__dirname, 'Student_Performance_Analytics_Presentation.pptx');
pptx.writeFile({ fileName: outputPath })
    .then(fileName => {
        console.log(`Minimal presentation generated successfully: ${fileName}`);
    })
    .catch(err => {
        console.error('Error generating presentation:', err);
        process.exit(1);
    });
