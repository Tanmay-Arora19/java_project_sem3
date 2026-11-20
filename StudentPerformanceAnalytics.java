import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class StudentPerformanceAnalytics {
    private static final Scanner input = new Scanner(System.in);
    private static final ArrayList<Student> students = new ArrayList<Student>();
    private static final ArrayList<String> subjects = new ArrayList<String>();
    private static final String DATA_FILE = "students.txt";

    interface Reportable {
        void printReport();
    }

    static class InvalidMarksException extends Exception {
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

        int getRollNumber() { return rollNumber; }
        String getName() { return name; }
        int getMarks(String subject) { return marks.get(subject); }

        int getTotalMarks() {
            int total = 0;
            for (int value : marks.values()) total += value;
            return total;
        }

        double getPercentage() {
            if (marks.isEmpty()) return 0.0;
            return getTotalMarks() * 100.0 / (marks.size() * 100.0);
        }

        String getGrade() {
            double p = getPercentage();
            if (p >= 90) return "A+";
            if (p >= 80) return "A";
            if (p >= 70) return "B";
            if (p >= 60) return "C";
            if (p >= 50) return "D";
            return "F";
        }

        String getHighestSubject() {
            String result = "-";
            int highest = -1;
            for (String subject : marks.keySet()) {
                if (marks.get(subject) > highest) {
                    highest = marks.get(subject);
                    result = subject;
                }
            }
            return result;
        }

        String getLowestSubject() {
            String result = "-";
            int lowest = 101;
            for (String subject : marks.keySet()) {
                if (marks.get(subject) < lowest) {
                    lowest = marks.get(subject);
                    result = subject;
                }
            }
            return result;
        }

        public void printReport() {
            System.out.println("\n---------- STUDENT PERFORMANCE REPORT ----------");
            System.out.println("Roll Number : " + rollNumber);
            System.out.println("Student Name: " + name);
            for (String subject : subjects) {
                if (marks.containsKey(subject))
                    System.out.println(subject + ": " + marks.get(subject) + "/100");
            }
            System.out.println("Total Marks : " + getTotalMarks() + "/" + (marks.size() * 100));
            System.out.printf("Percentage  : %.2f%%%n", getPercentage());
            System.out.println("Grade       : " + getGrade());
            System.out.println("Highest Subject: " + getHighestSubject());
            System.out.println("Lowest Subject : " + getLowestSubject());
            System.out.println("-------------------------------------------------");
        }

        public String toString() {
            return rollNumber + " - " + name + " | "
                    + String.format("%.2f", getPercentage())
                    + "% | Grade: " + getGrade();
        }
    }

    public static void main(String[] args) {
        loadData();
        boolean running = true;

        while (running) {
            showMenu();
            int choice = readInt("Enter your choice: ");

            switch (choice) {
                case 1: setupSubjects(); break;
                case 2: addStudent(); break;
                case 3: viewAllStudents(); break;
                case 4: searchStudentReport(); break;
                case 5: showStudentRanking(); break;
                case 6: showSubjectStatistics(); break;
                case 7: showClassSummary(); break;
                case 8: saveData(); createBackup(); break;
                case 9:
                    saveData();
                    System.out.println("Data saved. Thank you!");
                    running = false;
                    break;
                default:
                    System.out.println("Invalid choice. Please select 1 to 9.");
            }
        }
    }

    private static void showMenu() {
        System.out.println("\n======= STUDENT PERFORMANCE ANALYTICS TOOL =======");
        System.out.println("1. Set up subjects");
        System.out.println("2. Add student and marks");
        System.out.println("3. View all students");
        System.out.println("4. Search student report");
        System.out.println("5. Rank students");
        System.out.println("6. Subject-wise statistics");
        System.out.println("7. Class performance summary");
        System.out.println("8. Save data and create backup");
        System.out.println("9. Save and exit");
        System.out.println("==================================================");
    }

    private static int readInt(String message) {
        while (true) {
            System.out.print(message);
            try {
                return Integer.parseInt(input.nextLine().trim());
            } catch (NumberFormatException e) {
                System.out.println("Please enter a whole number.");
            }
        }
    }

    private static void setupSubjects() {
        if (!students.isEmpty()) {
            System.out.println("Subjects cannot be changed after student records are added.");
            return;
        }

        int count;
        do {
            count = readInt("How many subjects? ");
            if (count <= 0) System.out.println("Enter at least one subject.");
        } while (count <= 0);

        subjects.clear();
        for (int i = 0; i < count; i++) {
            System.out.print("Enter subject " + (i + 1) + " name: ");
            String subject = input.nextLine().trim();

            while (subject.isEmpty() || containsSubject(subject)) {
                System.out.print("Name is empty or already exists. Enter another: ");
                subject = input.nextLine().trim();
            }
            subjects.add(subject);
        }
        saveData();
        System.out.println("Subjects saved successfully.");
    }

    private static boolean containsSubject(String subject) {
        for (String old : subjects)
            if (old.equalsIgnoreCase(subject)) return true;
        return false;
    }

    private static void addStudent() {
        if (subjects.isEmpty()) {
            System.out.println("First select option 1 and add subjects.");
            return;
        }

        int roll = readInt("Enter student roll number: ");
        if (findStudent(roll) != null) {
            System.out.println("This roll number already exists.");
            return;
        }

        System.out.print("Enter student name: ");
        String name = input.nextLine().trim();
        while (name.isEmpty()) {
            System.out.print("Name cannot be empty. Enter again: ");
            name = input.nextLine().trim();
        }

        HashMap<String, Integer> studentMarks = new HashMap<String, Integer>();
        for (String subject : subjects)
            studentMarks.put(subject, readMarks(subject));

        Student student = new Student(roll, name, studentMarks);
        students.add(student);
        System.out.println("Student record added successfully.");
        student.printReport();

        final String savedName = name;
        Thread notice = new Thread(() ->
                System.out.println("[Notice] Marks record added for " + savedName + "."));
        notice.start();
        saveData();
    }

    private static int readMarks(String subject) {
        while (true) {
            int mark = readInt("Enter marks in " + subject + " (0-100): ");
            try {
                if (mark < 0 || mark > 100)
                    throw new InvalidMarksException("Marks must be between 0 and 100.");
                return mark;
            } catch (InvalidMarksException e) {
                System.out.println(e.getMessage());
            }
        }
    }

    private static void viewAllStudents() {
        if (students.isEmpty()) {
            System.out.println("No student records are available.");
            return;
        }
        System.out.println("\nROLL NO - NAME | PERCENTAGE | GRADE");
        displayList(students);
    }

    private static <T> void displayList(List<T> list) {
        for (T item : list) System.out.println(item);
    }

    private static Student findStudent(int roll) {
        for (Student student : students)
            if (student.getRollNumber() == roll) return student;
        return null;
    }

    private static void searchStudentReport() {
        int roll = readInt("Enter roll number to search: ");
        Student student = findStudent(roll);
        if (student == null) System.out.println("Student record not found.");
        else student.printReport();
    }

    private static void showStudentRanking() {
        if (students.isEmpty()) {
            System.out.println("No student records available for ranking.");
            return;
        }

        ArrayList<Student> ranked = new ArrayList<Student>(students);
        Comparator<Student> comparator =
                Comparator.comparingDouble(Student::getPercentage)
                        .reversed().thenComparing(Student::getName);
        Collections.sort(ranked, comparator);

        System.out.println("\n============= STUDENT RANKING =============");
        for (int i = 0; i < ranked.size(); i++) {
            Student s = ranked.get(i);
            System.out.printf("Rank %d: %s | %.2f%% | Grade %s%n",
                    i + 1, s.getName(), s.getPercentage(), s.getGrade());
        }
    }

    private static void showSubjectStatistics() {
        if (students.isEmpty()) {
            System.out.println("No student records available.");
            return;
        }

        System.out.println("\n========== SUBJECT-WISE STATISTICS ==========");
        for (String subject : subjects) {
            int total = 0, highest = -1, lowest = 101;
            String topStudent = "", lowStudent = "";

            for (Student s : students) {
                int mark = s.getMarks(subject);
                total += mark;
                if (mark > highest) {
                    highest = mark;
                    topStudent = s.getName();
                }
                if (mark < lowest) {
                    lowest = mark;
                    lowStudent = s.getName();
                }
            }

            System.out.printf("%nSubject: %s%n", subject);
            System.out.printf("Average Marks: %.2f/100%n", (double) total / students.size());
            System.out.println("Highest Marks: " + highest + " (by " + topStudent + ")");
            System.out.println("Lowest Marks : " + lowest + " (by " + lowStudent + ")");
        }
    }

    private static void showClassSummary() {
        if (students.isEmpty()) {
            System.out.println("No student records available.");
            return;
        }

        double average = students.stream()
                .mapToDouble(Student::getPercentage).average().orElse(0.0);
        Student topper = students.stream()
                .max(Comparator.comparingDouble(Student::getPercentage)).get();
        long passed = students.stream()
                .filter(s -> !s.getGrade().equals("F")).count();

        System.out.println("\n========== CLASS PERFORMANCE SUMMARY ==========");
        System.out.println("Total Students : " + students.size());
        System.out.printf("Class Average  : %.2f%%%n", average);
        System.out.println("Students Passed: " + passed);
        System.out.println("Students Failed: " + (students.size() - passed));
        System.out.println("Class Topper   : " + topper.getName());
        System.out.printf("Topper Percent : %.2f%%%n", topper.getPercentage());

        Map<String, Long> gradeCounts = students.stream()
                .collect(Collectors.groupingBy(
                        Student::getGrade, TreeMap::new, Collectors.counting()));
        System.out.println("Grade Distribution: " + gradeCounts);
    }

    private static void saveData() {
        try (PrintWriter writer = new PrintWriter(new FileWriter(DATA_FILE))) {
            writer.print("SUBJECTS");
            for (String subject : subjects) writer.print("|" + clean(subject));
            writer.println();

            for (Student s : students) {
                writer.print("STUDENT|" + s.getRollNumber() + "|" + clean(s.getName()));
                for (String subject : subjects)
                    writer.print("|" + s.getMarks(subject));
                writer.println();
            }
            System.out.println("Data saved in " + DATA_FILE);
        } catch (IOException e) {
            System.out.println("Could not save data: " + e.getMessage());
        }
    }

    private static void loadData() {
        File file = new File(DATA_FILE);
        if (!file.exists()) return;

        try (BufferedReader reader = new BufferedReader(new FileReader(file))) {
            String line = reader.readLine();
            if (line != null && line.startsWith("SUBJECTS")) {
                String[] parts = line.split("\\|", -1);
                for (int i = 1; i < parts.length; i++) subjects.add(parts[i]);
            }

            while ((line = reader.readLine()) != null) {
                String[] parts = line.split("\\|", -1);
                if (parts.length != subjects.size() + 3
                        || !parts[0].equals("STUDENT")) continue;

                try {
                    int roll = Integer.parseInt(parts[1]);
                    String name = parts[2];
                    HashMap<String, Integer> marks = new HashMap<String, Integer>();
                    for (int i = 0; i < subjects.size(); i++)
                        marks.put(subjects.get(i), Integer.parseInt(parts[i + 3]));
                    students.add(new Student(roll, name, marks));
                } catch (NumberFormatException e) {
                    System.out.println("Skipped one damaged student record.");
                }
            }
            System.out.println("Previous student records loaded.");
        } catch (IOException e) {
            System.out.println("Could not load data: " + e.getMessage());
        }
    }

    private static String clean(String value) {
        return value.replace("|", "/").replace("\n", " ");
    }

    private static void createBackup() {
        try (FileInputStream source = new FileInputStream(DATA_FILE);
             FileOutputStream destination =
                     new FileOutputStream("students_backup.txt")) {
            int data;
            while ((data = source.read()) != -1) destination.write(data);
            System.out.println("Backup created as students_backup.txt");
        } catch (IOException e) {
            System.out.println("Backup could not be created: " + e.getMessage());
        }
    }
}
