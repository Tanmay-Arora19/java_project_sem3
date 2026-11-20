const fs = require('fs');
const path = require('path');
const PptxGenJS = require(process.env.APPDATA + '/npm/node_modules/pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9'; // 10 x 5.625 inches
pptx.author = 'Tanmay Arora, Vishal, Jaskaran Singh, Ansh Arora';
pptx.company = 'Chitkara University';
pptx.title = 'Student Performance Analytics Tool';

// High-Impact Academic & Modern Palette
const NAVY = '0A192F';         // Deep Midnight Navy
const ROYAL = '134074';        // Slate / Royal Blue
const ACCENT = '0077B6';       // Academic Accent Blue
const SKY = '38BDF8';          // Electric Sky Blue Accent
const LIGHT_BG = 'F8FAFC';     // Pristine Off-White Canvas
const CARD_BG = 'FFFFFF';      // Pure White Card Fill
const DARK_TEXT = '0F172A';    // Charcoal Black Heading
const BODY_TEXT = '334155';    // Readable Slate
const MUTED_TEXT = '64748B';   // Muted Subtitle
const BORDER_COLOR = 'CBD5E1'; // Sleek Card Border
const GREEN = '166534';        // Emerald Success
const GREEN_BG = 'DCFCE7';
const RED = '991B1B';          // Crimson Alert
const RED_BG = 'FEE2E2';
const WHITE = 'FFFFFF';

// Helper function to create sleek header & footer
function addHeader(slide, titleText, subtitleText, slideNumber) {
    // Top banner
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 10, h: 0.82,
        fill: { color: NAVY }
    });
    // Accent electric stripe
    slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0.82, w: 10, h: 0.04,
        fill: { color: SKY }
    });
    // Title
    slide.addText(titleText, {
        x: 0.6, y: 0.1, w: 7.8, h: 0.4,
        fontSize: 18, fontFace: 'Arial', bold: true, color: WHITE
    });
    // Subtitle
    if (subtitleText) {
        slide.addText(subtitleText, {
            x: 0.6, y: 0.48, w: 7.8, h: 0.25,
            fontSize: 10.5, fontFace: 'Arial', color: '93C5FD'
        });
    }
    // Slide Number badge
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 8.8, y: 0.18, w: 0.65, h: 0.44,
        fill: { color: ROYAL },
        line: { color: SKY, width: 1 },
        rectRadius: 0.06
    });
    slide.addText(slideNumber.toString(), {
        x: 8.8, y: 0.18, w: 0.65, h: 0.44,
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
        fontSize: 8.5, fontFace: 'Arial', color: MUTED_TEXT
    });
    slide.addText('Student Performance Analytics Tool', {
        x: 5.6, y: 5.37, w: 3.8, h: 0.22,
        fontSize: 8.5, fontFace: 'Arial', color: ROYAL, align: 'right', bold: true
    });
}

// =============================================================
// SLIDE 1: TITLE PAGE (Impressive Executive Cover)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: NAVY };

    // Modern vertical accent stripes
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: 5.625, fill: { color: ACCENT } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.22, y: 0, w: 0.07, h: 5.625, fill: { color: SKY } });

    // Top Category Badge
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 0.45, w: 5.4, h: 0.38,
        fill: { color: ROYAL }, line: { color: SKY, width: 1 }, rectRadius: 0.06
    });
    slide.addText('B.TECH CSE (AI & ML) • 3RD SEMESTER COLLEGE PROJECT', {
        x: 0.8, y: 0.45, w: 5.4, h: 0.38,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: 'BAE6FD', align: 'center', valign: 'middle'
    });

    // Main Title
    slide.addText('STUDENT PERFORMANCE\nANALYTICS TOOL', {
        x: 0.8, y: 0.98, w: 8.5, h: 1.25,
        fontSize: 29, fontFace: 'Arial', bold: true, color: WHITE, lineSpacing: 36
    });

    // Subtitle
    slide.addText('A Terminal-Based Academic Record & Performance Evaluation System in Java', {
        x: 0.8, y: 2.34, w: 8.5, h: 0.35,
        fontSize: 12.5, fontFace: 'Arial', color: '93C5FD'
    });

    // Accent line
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.76, w: 8.4, h: 0.02, fill: { color: SKY } });

    // Left Card: Academic Details & Faculty Guide
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.8, y: 2.96, w: 4.05, h: 2.25,
        fill: { color: '0F172A' }, line: { color: '334155', width: 1 }, rectRadius: 0.08
    });
    slide.addText('ACADEMIC DETAILS', {
        x: 1.0, y: 3.12, w: 3.65, h: 0.25,
        fontSize: 11.5, fontFace: 'Arial', bold: true, color: SKY
    });
    const acadLines = [
        'Subject: Java Programming',
        'Course: B.Tech CSE (AI & ML)',
        'Institution: Chitkara University, Punjab',
        '',
        'Faculty Guide: Rohit Sharma'
    ];
    slide.addText(acadLines.join('\n'), {
        x: 1.0, y: 3.42, w: 3.65, h: 1.7,
        fontSize: 11, fontFace: 'Arial', color: WHITE, lineSpacing: 20
    });

    // Right Card: Project Team Members (Exact Order)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.15, y: 2.96, w: 4.05, h: 2.25,
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
// SLIDE 2: INTRODUCTION (Modern 4-Pillar Architecture)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Introduction', 'System Purpose & Core Pillars', 2);

    // Punchy Value Proposition Banner
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.0, w: 8.8, h: 0.6,
        fill: { color: CARD_BG }, line: { color: 'BAE6FD', width: 1 }, rectRadius: 0.06
    });
    slide.addText('A terminal-based Core Java application that eliminates manual grading errors, calculates instant rankings, and analyzes class performance in real time.', {
        x: 0.8, y: 1.05, w: 8.4, h: 0.5,
        fontSize: 11.5, fontFace: 'Arial', color: DARK_TEXT, align: 'center', valign: 'middle'
    });

    // 4 Modern Pillar Cards
    const pillars = [
        {
            title: '1. RECORD MANAGEMENT',
            points: [
                'Dynamic Subjects: Configure custom subject catalogs.',
                'Student Profiling: Stores roll numbers and names safely.',
                'Clean Collections: In-memory ArrayList & HashMap storage.'
            ],
            x: 0.6, y: 1.78
        },
        {
            title: '2. AUTOMATIC EVALUATION',
            points: [
                'Instant Math: Computes total marks & exact percentages.',
                'Grade Automation: Auto-assigns letter grades from A+ to F.',
                'Score Extremes: Identifies highest & lowest scoring subjects.'
            ],
            x: 5.1, y: 1.78
        },
        {
            title: '3. MERIT RANKING',
            points: [
                'Objective Ranking: Evaluates entire batch by percentage.',
                'Java Comparator: Automated descending sort logic.',
                'Fair Tie-Breaker: Sorts score ties alphabetically by name.'
            ],
            x: 0.6, y: 3.45
        },
        {
            title: '4. ACADEMIC ANALYTICS',
            points: [
                'Subject Insights: Generates averages, toppers, and lowest marks.',
                'Class Summary: Computes pass/fail counts & class topper.',
                'Permanent Storage: Saves records with automated byte backup.'
            ],
            x: 5.1, y: 3.45
        }
    ];

    pillars.forEach(p => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: p.x, y: p.y, w: 4.3, h: 1.55,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: p.x, y: p.y, w: 4.3, h: 0.34,
            fill: { color: ROYAL }
        });
        slide.addText(p.title, {
            x: p.x + 0.15, y: p.y + 0.05, w: 4.0, h: 0.24,
            fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(p.points.map(pt => '• ' + pt).join('\n'), {
            x: p.x + 0.15, y: p.y + 0.4, w: 4.0, h: 1.05,
            fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 18
        });
    });
}

// =============================================================
// SLIDE 3: PROBLEM STATEMENT (Visual Side-by-Side Comparison)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Problem Statement', 'Manual Grading Bottlenecks vs. Automated Java Solution', 3);

    // Left Column: Manual Challenges (Red Accent)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: 'FCA5A5', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: RED }
    });
    slide.addText('MANUAL EVALUATION ISSUES', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const issues = [
        'Time-Consuming: Hours spent calculating marks and percentages by hand for large classes.',
        'Prone to Errors: Manual calculation mistakes cause ranking errors and student disputes.',
        'Difficult Comparison: Comparing multi-subject scores on paper is confusing and slow.',
        'Data Loss Risk: Physical paper registers can easily get damaged, lost, or misplaced.',
        'Tedious Reporting: Generating batch averages, pass rates, and toppers takes extra manual work.'
    ];

    slide.addText(issues.map(i => '❌  ' + i).join('\n\n'), {
        x: 0.8, y: 1.6, w: 3.9, h: 3.3,
        fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 16
    });

    // Right Column: Our Java Solution (Green Accent)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: '86EFAC', width: 1.5 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: GREEN }
    });
    slide.addText('OUR AUTOMATED JAVA SOLUTION', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const solutions = [
        'Instant Automation: Real-time calculation of total marks, percentages, and letter grades.',
        '100% Accurate Ranking: Java Comparator guarantees fair, automated student sorting.',
        'Actionable Insights: Displays subject averages, top scorers, and lowest marks in seconds.',
        'Permanent Security: Files saved in students.txt with automated byte-stream backups.',
        'Effortless Navigation: Interactive numbered menu guides teachers step-by-step smoothly.'
    ];

    slide.addText(solutions.map(s => '✔  ' + s).join('\n\n'), {
        x: 5.3, y: 1.6, w: 3.9, h: 3.3,
        fontSize: 10, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 16
    });
}

// =============================================================
// SLIDE 4: OBJECTIVES OF THE PROJECT (Modern 6-Card Grid)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Project Objectives', 'Core System Capabilities & Deliverables', 4);

    const goals = [
        { num: '01', title: 'Dynamic Subject Setup', desc: 'Configure custom subject lists and accept marks dynamically for any class.' },
        { num: '02', title: 'Automated Percentage', desc: 'Compute exact percentages instantly with standard mathematical precision.' },
        { num: '03', title: 'Standard Letter Grading', desc: 'Automatically assign academic grades from A+ to F using threshold logic.' },
        { num: '04', title: 'Merit-Based Ranking', desc: 'Rank students objectively by percentage using Java Comparator sorting.' },
        { num: '05', title: 'Class & Subject Analytics', desc: 'Generate class averages, toppers, pass/fail rates, and subject breakdowns.' },
        { num: '06', title: 'File Persistence & Backup', desc: 'Save records permanently in text files with byte-stream disaster recovery.' }
    ];

    goals.forEach((g, idx) => {
        const col = idx % 2;
        const row = Math.floor(idx / 2);
        const x = col === 0 ? 0.6 : 5.1;
        const y = 1.05 + row * 1.35;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: y, w: 4.3, h: 1.2,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });

        // Left accent stripe
        slide.addShape(pptx.ShapeType.rect, {
            x: x, y: y, w: 0.12, h: 1.2,
            fill: { color: ROYAL }
        });

        // Number Pill
        slide.addShape(pptx.ShapeType.roundRect, {
            x: x + 0.25, y: y + 0.18, w: 0.5, h: 0.45,
            fill: { color: 'F0F7FF' }, line: { color: 'BAE6FD', width: 1 }, rectRadius: 0.04
        });
        slide.addText(g.num, {
            x: x + 0.25, y: y + 0.18, w: 0.5, h: 0.45,
            fontSize: 12, fontFace: 'Arial', bold: true, color: ACCENT, align: 'center', valign: 'middle'
        });

        // Title & Description
        slide.addText(g.title, {
            x: x + 0.85, y: y + 0.15, w: 3.3, h: 0.28,
            fontSize: 11, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(g.desc, {
            x: x + 0.85, y: y + 0.46, w: 3.3, h: 0.65,
            fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
        });
    });
}

// =============================================================
// SLIDE 5: TECHNOLOGIES AND JAVA CONCEPTS USED (Clean & Crisp)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Technologies & Java Concepts Used', 'Development Stack and OOP Implementation', 5);

    // Left Column: Tech Stack (Clean Card)
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 3.0, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 3.0, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('TECH SPECIFICATIONS', {
        x: 0.75, y: 1.1, w: 2.7, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const stackItems = [
        { label: 'Language', val: 'Java 8+ (Core Java)' },
        { label: 'Paradigm', val: 'Object-Oriented (OOP)' },
        { label: 'Interface', val: 'Console / Terminal CLI' },
        { label: 'Storage', val: 'Character Stream (students.txt)' },
        { label: 'Backup', val: 'Byte Stream (students_backup.txt)' },
        { label: 'IDE Used', val: 'VS Code / IntelliJ IDEA' },
        { label: 'Version Control', val: 'Git & GitHub' }
    ];

    let stkY = 1.55;
    stackItems.forEach(item => {
        slide.addText(item.label + ':', {
            x: 0.75, y: stkY, w: 2.7, h: 0.2,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: ACCENT
        });
        slide.addText(item.val, {
            x: 0.75, y: stkY + 0.2, w: 2.7, h: 0.26,
            fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT
        });
        stkY += 0.52;
    });

    // Right Column: Concepts Table
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 3.8, y: 1.02, w: 5.6, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 3.8, y: 1.02, w: 5.6, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('JAVA CONCEPT', { x: 3.95, y: 1.1, w: 2.1, h: 0.25, fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('HOW IT IS USED IN CODE', { x: 6.1, y: 1.1, w: 3.2, h: 0.25, fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE });

    const concepts = [
        { c: 'Classes & Objects', u: 'To model real-world student entities and marks.' },
        { c: 'Encapsulation', u: 'To protect student state using private variables & getters.' },
        { c: 'ArrayList & HashMap', u: 'Dynamic lists for students; Key-value marks mapping.' },
        { c: 'Reportable (Interface)', u: 'Defines a uniform contract for printing student reports.' },
        { c: 'InvalidMarksException', u: 'Custom exception validating marks strictly between 0–100.' },
        { c: 'Comparator', u: 'Ranks students by percentage with alphabetical tie-breaking.' },
        { c: 'Stream API', u: 'Calculates class average, pass/fail counts, and topper.' },
        { c: 'Generics', u: 'Reusable displayList() method for printing lists of any type.' },
        { c: 'Multithreading', u: 'Background notice thread when new records are saved.' },
        { c: 'File I/O Streams', u: 'Character streams for text data; Byte streams for backup.' }
    ];

    let rowY = 1.48;
    concepts.forEach((row, i) => {
        const bg = (i % 2 === 0) ? 'FFFFFF' : 'F8FAFC';
        slide.addShape(pptx.ShapeType.rect, {
            x: 3.85, y: rowY, w: 5.5, h: 0.35,
            fill: { color: bg }
        });
        slide.addText(row.c, {
            x: 3.95, y: rowY + 0.05, w: 2.1, h: 0.25,
            fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY
        });
        slide.addText(row.u, {
            x: 6.1, y: rowY + 0.05, w: 3.2, h: 0.25,
            fontSize: 9, fontFace: 'Arial', color: BODY_TEXT
        });
        rowY += 0.36;
    });
}

// =============================================================
// SLIDE 6: PROJECT ARCHITECTURE (Clean Modern Flowchart)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Project Architecture', 'System Workflow from User Input to File Persistence', 6);

    const steps = [
        { label: 'TEACHER USER\nTerminal Interaction', color: NAVY },
        { label: 'MAIN MENU\nSwitch-Case Routing', color: ROYAL },
        { label: 'INPUT VALIDATION\nDuplicate & 0-100 Check', color: ACCENT },
        { label: 'STUDENT MODEL\nEncapsulated Object', color: ROYAL },
        { label: 'CALCULATION ENGINE\nTotals, % & Grades', color: NAVY },
        { label: 'RANKING & ANALYTICS\nComparator & Streams', color: ROYAL },
        { label: 'REPORT GENERATOR\nReportable Interface', color: ACCENT },
        { label: 'FILE & BACKUP I/O\nstudents.txt & Backup', color: NAVY }
    ];

    // Arrange in 2 rows of 4 boxes with connecting arrows
    steps.forEach((st, idx) => {
        const isRow1 = idx < 4;
        const col = isRow1 ? idx : (7 - idx);
        const x = 0.6 + col * 2.25;
        const y = isRow1 ? 1.15 : 2.75;

        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: y, w: 2.05, h: 1.0,
            fill: { color: st.color },
            line: { color: SKY, width: 1 },
            rectRadius: 0.08
        });
        slide.addText(st.label, {
            x: x + 0.08, y: y + 0.14, w: 1.89, h: 0.72,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle'
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

    // Bottom Architecture Summary
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 4.1, w: 8.8, h: 1.05,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Key Architectural Takeaway:', {
        x: 0.8, y: 4.22, w: 8.4, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: NAVY
    });
    slide.addText('Modular separation of concerns: The Main Menu dispatches commands, the Student class encapsulates data safely, the Analytics engine calculates results, and File I/O ensures persistent storage with automated backup.', {
        x: 0.8, y: 4.46, w: 8.4, h: 0.58,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 7: STUDENT REGISTRATION AND DATA MANAGEMENT
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Student Registration & Data Management', 'Data Collection, Integrity Checks, and In-Memory Storage', 7);

    const cards = [
        {
            title: '1. DATA INPUTS',
            color: NAVY,
            points: [
                'Roll Number: Unique numerical identifier.',
                'Student Name: Validated non-empty string.',
                'Subject Catalog: Configured dynamically once.',
                'Marks: Numerical score per subject (0–100).'
            ],
            x: 0.6
        },
        {
            title: '2. VALIDATION RULES',
            color: GREEN,
            points: [
                'Duplicate Check: Scans & rejects existing roll numbers.',
                'Custom Exception: InvalidMarksException (0–100 range).',
                'Safe Re-prompt: Re-prompts cleanly without crashing.',
                'Setup Gate: Enforces subject setup before student entry.'
            ],
            x: 3.6
        },
        {
            title: '3. JAVA COLLECTIONS',
            color: ROYAL,
            points: [
                'ArrayList<Student>: Stores student objects dynamically.',
                'HashMap<String, Integer>: Maps subject names to marks.',
                'ArrayList<String>: Maintains active subject list.',
                'Fast Access: Instant lookups and fast iteration.'
            ],
            x: 6.6
        }
    ];

    cards.forEach(c => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: c.x, y: 1.05, w: 2.8, h: 2.9,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addShape(pptx.ShapeType.rect, {
            x: c.x, y: 1.05, w: 2.8, h: 0.38,
            fill: { color: c.color }
        });
        slide.addText(c.title, {
            x: c.x + 0.1, y: 1.12, w: 2.6, h: 0.24,
            fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE
        });
        slide.addText(c.points.map(p => '• ' + p).join('\n\n'), {
            x: c.x + 0.12, y: 1.52, w: 2.56, h: 2.3,
            fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
        });
    });

    // Bottom Summary Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 4.15, w: 8.8, h: 1.0,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addText('Viva Takeaway:', {
        x: 0.8, y: 4.25, w: 8.4, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: NAVY
    });
    slide.addText('"Data integrity is strictly enforced at entry time. Invalid marks trigger InvalidMarksException, duplicate rolls are rejected, and marks are linked via HashMaps for optimal organization."', {
        x: 0.8, y: 4.48, w: 8.4, h: 0.55,
        fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT, italic: true
    });
}

// =============================================================
// SLIDE 8: PERCENTAGE AND GRADE CALCULATION (Sleek Visuals)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Percentage & Grade Calculation', 'Formulas, Step-by-Step Example, and Grading Scale', 8);

    // Left Column: Formula & Example
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
        x: 0.8, y: 1.55, w: 3.9, h: 0.75,
        fill: { color: 'F0F7FF' }, line: { color: 'BAE6FD', width: 1 }, rectRadius: 0.06
    });
    slide.addText('CALCULATION FORMULA:', {
        x: 0.9, y: 1.62, w: 3.7, h: 0.2,
        fontSize: 9, fontFace: 'Arial', bold: true, color: ACCENT
    });
    slide.addText('Percentage = (Total Marks Obtained / Maximum Total Marks) × 100', {
        x: 0.9, y: 1.88, w: 3.7, h: 0.35,
        fontSize: 9.5, fontFace: 'Arial', bold: true, color: NAVY
    });

    // Practical Example
    slide.addText('WORKED EXAMPLE (3 SUBJECTS):', {
        x: 0.8, y: 2.42, w: 3.9, h: 0.22,
        fontSize: 10, fontFace: 'Arial', bold: true, color: DARK_TEXT
    });
    const exampleLines = [
        '• Java Programming   = 85 / 100',
        '• Mathematics III    = 90 / 100',
        '• Data Structures    = 75 / 100',
        '--------------------------------------------',
        'Total Marks Obtained  = 85 + 90 + 75 = 250',
        'Maximum Total Marks   = 3 × 100 = 300',
        'Percentage = (250 / 300) × 100 = 83.33%',
        'Resulting Grade = A (Excellent)'
    ];
    slide.addText(exampleLines.join('\n'), {
        x: 0.8, y: 2.7, w: 3.9, h: 2.3,
        fontSize: 9.5, fontFace: 'Courier New', color: BODY_TEXT, lineSpacing: 15
    });

    // Right Column: Official Grading Scale
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('ACADEMIC GRADING SCALE', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    // Table Header
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.2, y: 1.55, w: 4.1, h: 0.32,
        fill: { color: 'E2E8F0' }
    });
    slide.addText('PERCENTAGE', { x: 5.3, y: 1.62, w: 1.6, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY });
    slide.addText('GRADE', { x: 7.0, y: 1.62, w: 0.8, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY });
    slide.addText('PERFORMANCE', { x: 7.9, y: 1.62, w: 1.3, h: 0.2, fontSize: 9, fontFace: 'Arial', bold: true, color: NAVY });

    const grades = [
        { r: '90% – 100%', g: 'A+', rem: 'Outstanding', col: GREEN },
        { r: '80% – 89%',  g: 'A',  rem: 'Excellent',   col: GREEN },
        { r: '70% – 79%',  g: 'B',  rem: 'Very Good',   col: NAVY },
        { r: '60% – 69%',  g: 'C',  rem: 'Good',        col: NAVY },
        { r: '50% – 59%',  g: 'D',  rem: 'Pass',        col: 'B45309' },
        { r: 'Below 50%', g: 'F',  rem: 'Fail',        col: RED }
    ];

    let gY = 1.9;
    grades.forEach(gr => {
        slide.addText(gr.r, { x: 5.3, y: gY, w: 1.6, h: 0.24, fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT });
        slide.addText(gr.g, { x: 7.0, y: gY, w: 0.8, h: 0.24, fontSize: 10.5, fontFace: 'Arial', bold: true, color: gr.col });
        slide.addText(gr.rem, { x: 7.9, y: gY, w: 1.3, h: 0.24, fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT });
        gY += 0.38;
    });

    slide.addText('Implementation Note:\nThe getGrade() method uses conditional if-else statements to automatically assign the letter grade immediately after computing percentage.', {
        x: 5.3, y: 4.3, w: 3.9, h: 0.75,
        fontSize: 9, fontFace: 'Arial', color: MUTED_TEXT, lineSpacing: 14
    });
}

// =============================================================
// SLIDE 9: STUDENT RANKING USING COMPARATOR (Visual Leaderboard)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Student Ranking Using Comparator', 'Automated Sorting and Tie-Breaking Logic', 9);

    // Left Column: How Ranking Works
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
        'Objective Evaluation: Compares overall percentages across all students.',
        'Descending Order: Highest percentage receives Rank 1 automatically.',
        'Java Comparator: Chained comparator ensures clean, declarative sorting.',
        'Tie-Breaker: When percentages match, students are sorted alphabetically by name.',
        'Non-Destructive: Sorts a cloned ArrayList so original entry order is preserved.'
    ];

    slide.addText(rankPoints.map(p => '• ' + p).join('\n\n'), {
        x: 0.8, y: 1.55, w: 3.9, h: 2.3,
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
    slide.addText('SAMPLE MERIT LEADERBOARD', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const rankSample = [
        { rank: 'Rank 1', name: 'Tanmay Arora', pct: '92.00%', grade: 'A+', badge: 'TOPPER', bg: 'DCFCE7', col: GREEN },
        { rank: 'Rank 2', name: 'Vishal', pct: '88.33%', grade: 'A', badge: 'EXCELLENT', bg: 'F0F7FF', col: NAVY },
        { rank: 'Rank 3', name: 'Jaskaran Singh', pct: '85.67%', grade: 'A', badge: 'EXCELLENT', bg: 'FFFFFF', col: NAVY },
        { rank: 'Rank 4', name: 'Ansh Arora', pct: '81.00%', grade: 'A', badge: 'EXCELLENT', bg: 'F8FAFC', col: NAVY },
        { rank: 'Rank 5', name: 'Rohan Verma', pct: '74.50%', grade: 'B', badge: 'VERY GOOD', bg: 'FFFFFF', col: BODY_TEXT }
    ];

    let rY = 1.6;
    rankSample.forEach(r => {
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.25, y: rY, w: 4.0, h: 0.44,
            fill: { color: r.bg }, line: { color: BORDER_COLOR, width: 0.5 }, rectRadius: 0.04
        });
        slide.addText(r.rank, { x: 5.35, y: rY + 0.1, w: 0.8, h: 0.24, fontSize: 9.5, fontFace: 'Arial', bold: true, color: r.col });
        slide.addText(r.name, { x: 6.15, y: rY + 0.1, w: 1.4, h: 0.24, fontSize: 9.5, fontFace: 'Arial', bold: true, color: DARK_TEXT });
        slide.addText(r.pct + ' (' + r.grade + ')', { x: 7.55, y: rY + 0.1, w: 1.0, h: 0.24, fontSize: 9, fontFace: 'Arial', color: BODY_TEXT });
        slide.addShape(pptx.ShapeType.rect, { x: 8.55, y: rY + 0.1, w: 0.6, h: 0.24, fill: { color: r.col } });
        slide.addText(r.badge, { x: 8.55, y: rY + 0.1, w: 0.6, h: 0.24, fontSize: 7, fontFace: 'Arial', bold: true, color: WHITE, align: 'center', valign: 'middle' });
        rY += 0.5;
    });

    slide.addText('Takeaway: Eliminates human bias and hours of manual merit compilation.', {
        x: 5.3, y: 4.3, w: 3.9, h: 0.6,
        fontSize: 9.5, fontFace: 'Arial', color: NAVY, italic: true
    });
}

// =============================================================
// SLIDE 10: SUBJECT-WISE STATISTICS (Analytics Table & Insights)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Subject-Wise Statistics', 'Subject Difficulty, Class Average, and Top Scorers', 10);

    // Summary Table Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.05, w: 8.8, h: 2.6,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.05, w: 8.8, h: 0.4,
        fill: { color: NAVY }
    });

    slide.addText('SUBJECT NAME', { x: 0.8, y: 1.13, w: 2.2, h: 0.24, fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('CLASS AVERAGE', { x: 3.0, y: 1.13, w: 1.8, h: 0.24, fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('HIGHEST SCORE & TOPPER', { x: 4.8, y: 1.13, w: 2.4, h: 0.24, fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE });
    slide.addText('LOWEST SCORE', { x: 7.2, y: 1.13, w: 2.0, h: 0.24, fontSize: 10, fontFace: 'Arial', bold: true, color: WHITE });

    const subjectStats = [
        { sub: 'Java Programming', avg: '78.50 / 100', high: '92 / 100 (Aman Sharma)', low: '65 / 100 (Priya Verma)', bg: 'FFFFFF' },
        { sub: 'Mathematics III', avg: '81.00 / 100', high: '95 / 100 (Rahul Gupta)', low: '58 / 100 (Simran Kaur)', bg: 'F8FAFC' },
        { sub: 'Data Structures (DSA)', avg: '74.00 / 100', high: '88 / 100 (Neha Patel)', low: '60 / 100 (Rohan Mehta)', bg: 'FFFFFF' }
    ];

    let subY = 1.55;
    subjectStats.forEach(st => {
        slide.addShape(pptx.ShapeType.rect, {
            x: 0.65, y: subY, w: 8.7, h: 0.5,
            fill: { color: st.bg }
        });
        slide.addText(st.sub, { x: 0.8, y: subY + 0.12, w: 2.2, h: 0.26, fontSize: 10.5, fontFace: 'Arial', bold: true, color: NAVY });
        slide.addText(st.avg, { x: 3.0, y: subY + 0.12, w: 1.8, h: 0.26, fontSize: 10.5, fontFace: 'Arial', color: BODY_TEXT });
        slide.addText(st.high, { x: 4.8, y: subY + 0.12, w: 2.4, h: 0.26, fontSize: 10.5, fontFace: 'Arial', bold: true, color: GREEN });
        slide.addText(st.low, { x: 7.2, y: subY + 0.12, w: 2.0, h: 0.26, fontSize: 10.5, fontFace: 'Arial', color: RED });
        subY += 0.55;
    });

    // 3 Strategic Insight Cards
    const subBenefits = [
        {
            title: 'Identify Subject Strengths',
            desc: 'High average indicates topics well-mastered by the batch.'
        },
        {
            title: 'Detect Difficult Areas',
            desc: 'Lower averages alert teachers to schedule remedial sessions.'
        },
        {
            title: 'Recognize Top Talent',
            desc: 'Celebrates individual student brilliance in specific subjects.'
        }
    ];

    subBenefits.forEach((b, idx) => {
        const x = 0.6 + idx * 3.05;
        slide.addShape(pptx.ShapeType.roundRect, {
            x: x, y: 3.85, w: 2.8, h: 1.3,
            fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
        });
        slide.addText('✔ ' + b.title, {
            x: x + 0.15, y: 3.98, w: 2.5, h: 0.25,
            fontSize: 10, fontFace: 'Arial', bold: true, color: ROYAL
        });
        slide.addText(b.desc, {
            x: x + 0.15, y: 4.28, w: 2.5, h: 0.75,
            fontSize: 9, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
        });
    });
}

// =============================================================
// SLIDE 11: CLASS PERFORMANCE SUMMARY (Batch Metrics & Streams)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Class Performance Summary', 'Batch-Level Academic Insights via Java Stream API', 11);

    // Left Column: Key Metrics Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('OVERALL CLASS METRICS', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 11, fontFace: 'Arial', bold: true, color: WHITE
    });

    const metrics = [
        { label: 'Total Students', val: '5 Students in Sample Batch' },
        { label: 'Class Average', val: '84.31% Mean Score' },
        { label: 'Passed Students', val: '5 Students (100% Pass Rate)' },
        { label: 'Failed Students', val: '0 Students (grades below 50%)' },
        { label: 'Class Topper', val: 'Tanmay Arora (92.00%, Grade A+)' },
        { label: 'Grade Distribution', val: '{A+: 1, A: 3, B: 1, C: 0, D: 0, F: 0}' }
    ];

    let mY = 1.6;
    metrics.forEach(m => {
        slide.addText('• ' + m.label + ':', {
            x: 0.8, y: mY, w: 3.9, h: 0.2,
            fontSize: 9.5, fontFace: 'Arial', bold: true, color: ACCENT
        });
        slide.addText(m.val, {
            x: 0.95, y: mY + 0.2, w: 3.75, h: 0.25,
            fontSize: 9.5, fontFace: 'Arial', color: DARK_TEXT
        });
        mY += 0.52;
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
            goal: '1. Class Average',
            desc: 'students.stream().mapToDouble(Student::getPercentage).average() calculates the exact class mean percentage.'
        },
        {
            goal: '2. Class Topper',
            desc: 'students.stream().max(Comparator.comparingDouble(...)) identifies top student without manual loops.'
        },
        {
            goal: '3. Pass / Fail Counts',
            desc: 'students.stream().filter(s -> !s.getGrade().equals("F")).count() tallies passed students cleanly.'
        },
        {
            goal: '4. Grade Distribution',
            desc: 'Collectors.groupingBy(Student::getGrade, TreeMap::new, ...) groups student counts by grade letter.'
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

    slide.addText('Viva Takeaway: Stream API enables declarative data filtering and aggregation without cumbersome nested loops.', {
        x: 5.25, y: 4.3, w: 3.9, h: 0.65,
        fontSize: 9, fontFace: 'Arial', color: NAVY, italic: true
    });
}

// =============================================================
// SLIDE 12: FILE HANDLING AND BACKUP SYSTEM (Dual Stream Engine)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'File Handling & Backup System', 'Persistent Storage Architecture and Automated Disaster Recovery', 12);

    // Left Column: Character Streams
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('PRIMARY STORAGE (students.txt)', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const charStreamPoints = [
        'Storage Target: Permanent disk file named students.txt.',
        'Saving Engine: FileWriter & PrintWriter write formatted text rows.',
        'Loading Engine: FileReader & BufferedReader restore previous records.',
        'Pipe Format: SUBJECTS|Java|Math & STUDENT|101|Name|85|90...',
        'Persistence: Data remains available across application restarts.'
    ];

    slide.addText(charStreamPoints.map(p => '• ' + p).join('\n\n'), {
        x: 0.8, y: 1.6, w: 3.9, h: 3.3,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
    });

    // Right Column: Byte Streams
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.4,
        fill: { color: NAVY }
    });
    slide.addText('BACKUP ENGINE (students_backup.txt)', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const byteStreamPoints = [
        'Backup Target: Disaster recovery file students_backup.txt.',
        'Byte Cloning: FileInputStream & FileOutputStream duplicate raw bytes.',
        'On-Demand Trigger: Activated via Menu Option 8.',
        'Disaster Recovery: Protects data if primary file is corrupted or lost.',
        'Viva Distinction: Character streams handle text; Byte streams copy raw binary.'
    ];

    slide.addText(byteStreamPoints.map(p => '✔ ' + p).join('\n\n'), {
        x: 5.3, y: 1.6, w: 3.9, h: 3.3,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 15
    });
}

// =============================================================
// SLIDE 13: DIVISION OF WORK AMONG TEAM MEMBERS (Exact Sequence)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Division of Work Among Team Members', 'Individual Module Ownership & Technical Contributions', 13);

    const teamWork = [
        {
            num: '1',
            name: 'VISHAL',
            role: 'Student Data Management Module',
            color: NAVY,
            tasks: [
                'Created the core Student class and private fields.',
                'Implemented constructors, accessors, and toString().',
                'Developed total marks and percentage calculations.',
                'Implemented academic letter grading logic (A+ to F).',
                'Identified highest & lowest scoring subjects.',
                'Implemented the Reportable interface contract.'
            ],
            x: 0.6, y: 1.02
        },
        {
            num: '2',
            name: 'JASKARAN SINGH',
            role: 'Registration & Record Validation',
            color: ROYAL,
            tasks: [
                'Built dynamic subject setup module (Option 1).',
                'Implemented student registration workflow (Option 2).',
                'Accepted subject marks with range validation.',
                'Implemented duplicate roll number prevention.',
                'Created custom InvalidMarksException (0–100 check).',
                'Implemented student search and generic displayList().'
            ],
            x: 5.1, y: 1.02
        },
        {
            num: '3',
            name: 'ANSH ARORA',
            role: 'Analytics & Ranking Engine',
            color: ROYAL,
            tasks: [
                'Implemented student ranking using Java Comparator.',
                'Built percentage sorting and name tie-breaker logic.',
                'Calculated subject-wise averages, highest, and lowest scores.',
                'Developed overall class average and topper detection.',
                'Calculated passing and failing student counts.',
                'Generated grade distribution map using Stream API.'
            ],
            x: 0.6, y: 3.15
        },
        {
            num: '4',
            name: 'TANMAY ARORA',
            role: 'Architecture, File I/O & Integration',
            color: NAVY,
            tasks: [
                'Designed interactive console menu and navigation.',
                'Integrated all individual modules into a unified tool.',
                'Built character-stream persistence (FileWriter & PrintWriter).',
                'Implemented automated record loading upon startup.',
                'Created byte-stream backup engine (FileInputStream & Output).',
                'Built notification worker thread and Git repository.'
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
// SLIDE 14: LIVE PROJECT DEMONSTRATION (10-Step Workflow + Output)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Live Project Demonstration', 'Execution Sequence & Sample Terminal Report Card', 14);

    // Left Column: 10 Steps
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.4, h: 4.15,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.4, h: 0.4,
        fill: { color: ROYAL }
    });
    slide.addText('10-STEP EXECUTION FLOW', {
        x: 0.75, y: 1.1, w: 4.1, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const steps = [
        'Step 1:  Launch tool — existing records load automatically.',
        'Step 2:  Select Option 1 to set up class subjects.',
        'Step 3:  Enter subjects (e.g. Java, Mathematics, DSA).',
        'Step 4:  Select Option 2 to register student and enter marks.',
        'Step 5:  Enter unique roll number and student name.',
        'Step 6:  Enter marks for each subject (0 to 100 validation).',
        'Step 7:  View generated student report card instantly.',
        'Step 8:  Select Option 5 to display ranked leaderboard.',
        'Step 9:  Select Option 6 & 7 for subject and class analytics.',
        'Step 10: Select Option 8 to save records and clone backup.'
    ];

    slide.addText(steps.join('\n'), {
        x: 0.8, y: 1.55, w: 4.05, h: 3.5,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 18
    });

    // Right Column: Terminal Output Preview
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
        x: 5.35, y: 1.55, w: 3.9, h: 3.5,
        fontSize: 9, fontFace: 'Courier New', color: 'E2E8F0', lineSpacing: 14
    });
}

// =============================================================
// SLIDE 15: CONCLUSION AND FUTURE SCOPE (Executive Wrap-Up)
// =============================================================
{
    const slide = pptx.addSlide();
    slide.background = { color: LIGHT_BG };
    addHeader(slide, 'Conclusion & Future Scope', 'Project Achievements & Future Roadmap', 15);

    // Left Column: Conclusion Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 1.02, w: 4.3, h: 2.9,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 0.6, y: 1.02, w: 4.3, h: 0.38,
        fill: { color: ROYAL }
    });
    slide.addText('PROJECT TAKEAWAYS', {
        x: 0.75, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const conclusions = [
        'Automated Efficiency: Replaces manual paper registers and tedious calculations.',
        'Zero Calculation Error: Eliminates human mistakes in percentages and rankings.',
        'Actionable Insights: Delivers instant subject and batch-level performance analytics.',
        'Permanent Data Security: Safeguards records via file persistence & byte backup.',
        'Core Java Mastery: Combines OOP, Collections, Generics, Streams, and Multithreading.'
    ];

    slide.addText(conclusions.map(c => '• ' + c).join('\n\n'), {
        x: 0.75, y: 1.5, w: 4.0, h: 2.3,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });

    // Right Column: Future Scope Card
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 5.1, y: 1.02, w: 4.3, h: 2.9,
        fill: { color: CARD_BG }, line: { color: BORDER_COLOR, width: 1 }, rectRadius: 0.08
    });
    slide.addShape(pptx.ShapeType.rect, {
        x: 5.1, y: 1.02, w: 4.3, h: 0.38,
        fill: { color: NAVY }
    });
    slide.addText('FUTURE ROADMAP', {
        x: 5.25, y: 1.1, w: 4.0, h: 0.25,
        fontSize: 10.5, fontFace: 'Arial', bold: true, color: WHITE
    });

    const futureScope = [
        'Graphical User Interface: Desktop GUI built with JavaFX or Swing.',
        'Database Integration: Relational database connectivity via MySQL & JDBC.',
        'Authentication Portal: Role-based access for teachers and students.',
        'Export Capabilities: Automated printable PDF report card generation.',
        'Visual Performance Charts: Graphical bar and pie charts for class trends.'
    ];

    slide.addText(futureScope.map(f => '✔ ' + f).join('\n\n'), {
        x: 5.25, y: 1.5, w: 4.0, h: 2.3,
        fontSize: 9.5, fontFace: 'Arial', color: BODY_TEXT, lineSpacing: 14
    });

    // Bottom Banner: Thank You!
    slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.6, y: 4.12, w: 8.8, h: 1.05,
        fill: { color: NAVY }, line: { color: SKY, width: 1 }, rectRadius: 0.08
    });
    slide.addText('THANK YOU!', {
        x: 0.8, y: 4.2, w: 8.4, h: 0.45,
        fontSize: 24, fontFace: 'Arial', bold: true, color: WHITE, align: 'center'
    });
    slide.addText('Any Questions? We would be glad to answer.', {
        x: 0.8, y: 4.68, w: 8.4, h: 0.35,
        fontSize: 12.5, fontFace: 'Arial', color: SKY, align: 'center'
    });
}

// Generate PPTX file
const outputPath = path.join(__dirname, 'Student_Performance_Analytics_Presentation.pptx');
pptx.writeFile({ fileName: outputPath })
    .then(fileName => {
        console.log(`Impressive presentation generated successfully: ${fileName}`);
    })
    .catch(err => {
        console.error('Error generating presentation:', err);
        process.exit(1);
    });
