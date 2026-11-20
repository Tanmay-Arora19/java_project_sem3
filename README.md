# Student Performance Analytics Tool

A modular, console-based Java application built to manage student academic records, calculate grading and rankings, compute subject and class statistics, and provide file persistence with automated backups.

Developed as a Semester 3 Core Java and Object-Oriented Programming (OOP) project.

---

## Key Features

- **Subject Configuration**: Configure custom subject lists dynamically for an academic batch.
- **Student Record Management**: Add students with roll number deduplication, name validation, and subject marks entry.
- **Custom Exception Handling**: Validates mark inputs using a custom `InvalidMarksException` ensuring scores stay strictly within 0–100.
- **Detailed Academic Reporting**: Evaluates total scores, percentage, letter grades (`A+`, `A`, `B`, `C`, `D`, `F`), and identifies each student's strongest and weakest subjects.
- **Academic Ranking**: Sorts and ranks all students by overall percentage in descending order with alphabetical tie-breaking.
- **Subject-Wise Analytics**: Calculates class averages, highest score (with top student), and lowest score for every registered subject.
- **Class Performance Summary**: Leverages the Java Streams API to aggregate class average percentage, passing and failing counts, class topper, and grade distribution.
- **Data Persistence & Byte Backup**: Automatically saves and loads records via character streams (`students.txt`) and generates byte-stream backup copies (`students_backup.txt`).
- **Concurrent Notification**: Dispatches background worker threads to notify the console when records are committed without interrupting data entry.

---

## Technical Highlights & OOP Principles

| OOP / Java Concept | Implementation Details |
| :--- | :--- |
| **Encapsulation** | `Student` class encapsulates private attributes (`rollNumber`, `name`, `marks`) exposed through getter and calculation methods. |
| **Abstraction & Interfaces** | `Reportable` interface specifies contract for report generation implemented by `Student`. |
| **Custom Exceptions** | `InvalidMarksException` handles user input validation errors for marks outside the valid range. |
| **Collections Framework** | `ArrayList` maintains subject and student sequences; `HashMap` maps subject names to marks; `TreeMap` sorts grade distributions. |
| **Generics** | Generic utility method `<T> void displayList(List<T> list)` handles printing of arbitrary element collections. |
| **Java Streams & Lambdas** | Stream pipelines (`filter`, `mapToDouble`, `Collectors.groupingBy`, `Collectors.counting`) generate summary analytics. |
| **Concurrency / Multithreading** | Spawns separate `Thread` execution for non-blocking console notifications upon student creation. |
| **File I/O (Persistence & Backup)** | Character streams (`PrintWriter`, `BufferedReader`) for formatted text data; Byte streams (`FileInputStream`, `FileOutputStream`) for backup cloning. |

---

## Project Structure

```text
java_project_sem3/
├── StudentPerformanceAnalytics.java   # Main application source code
├── README.md                          # Project documentation
├── .gitignore                         # Git ignore configuration
├── run.bat                            # Windows one-click build and run script
└── run.sh                             # Linux / macOS shell execution script
```

---

## Getting Started

### Prerequisites

- Java Development Kit (JDK) 8 or higher
- Git installed on your system

Verify your Java installation:
```bash
javac -version
java -version
```

---

### Compilation & Execution

#### Option 1: Using the helper scripts

- **Windows**:
  Double-click `run.bat` or run in Command Prompt / PowerShell:
  ```cmd
  run.bat
  ```

- **Linux / macOS**:
  ```bash
  chmod +x run.sh
  ./run.sh
  ```

#### Option 2: Using the command line directly

1. Compile the source file:
   ```bash
   javac StudentPerformanceAnalytics.java
   ```

2. Run the application:
   ```bash
   java StudentPerformanceAnalytics
   ```

---

## Application Menu

When started, the tool presents an interactive console menu:

```text
======= STUDENT PERFORMANCE ANALYTICS TOOL =======
1. Set up subjects
2. Add student and marks
3. View all students
4. Search student report
5. Rank students
6. Subject-wise statistics
7. Class performance summary
8. Save data and create backup
9. Save and exit
==================================================
```

### Grading Scale Reference

| Percentage Range | Grade | Description |
| :--- | :--- | :--- |
| **90.00% – 100%** | `A+` | Outstanding |
| **80.00% – 89.99%** | `A`  | Excellent |
| **70.00% – 79.99%** | `B`  | Very Good |
| **60.00% – 69.99%** | `C`  | Good |
| **50.00% – 59.99%** | `D`  | Pass |
| **Below 50.00%**    | `F`  | Fail |

---

## File Storage Format

Data is preserved in `students.txt` using a delimited format:
- **Header Line**: `SUBJECTS|<Subject1>|<Subject2>|...`
- **Student Record**: `STUDENT|<RollNumber>|<StudentName>|<Mark1>|<Mark2>|...`

An exact copy is maintained in `students_backup.txt` when option 8 is selected or when the backup routine is triggered.

---

## Author

Developed by **Tanmay Arora**  
GitHub: [@Tanmay-Arora19](https://github.com/Tanmay-Arora19)
