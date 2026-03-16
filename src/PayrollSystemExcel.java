import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PayrollSystemExcel {

    // ===================== DATA HOLDERS =====================
    static ArrayList<Integer> empNos = new ArrayList<>();
    static ArrayList<String> empNames = new ArrayList<>();
    static ArrayList<String> empBirthdays = new ArrayList<>();
    static ArrayList<Double> basicSalaries = new ArrayList<>();
    static ArrayList<Double> hourlyRates = new ArrayList<>();

    // Attendance: each entry = {employeeNumber, dayOfMonth, hoursWorked}
    static ArrayList<double[]> attendanceList = new ArrayList<>();

    // ===================== EXCEL READING =====================

    /**
     * Load employee details and attendance records from the Excel file.
     */
    static void loadExcelData(String filePath) throws IOException {
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);

        // --- Sheet 1: Employee Details ---
        Sheet empSheet = workbook.getSheet("Employee Details");
        for (int i = 1; i <= empSheet.getLastRowNum(); i++) {
            Row row = empSheet.getRow(i);
            if (row == null) continue;

            int empNo = (int) row.getCell(0).getNumericCellValue();
            String lastName = row.getCell(1).getStringCellValue().trim();
            String firstName = row.getCell(2).getStringCellValue().trim();

            // Birthday (column 3) - stored as numeric (Excel serial date)
            Cell bdayCell = row.getCell(3);
            String birthday = "";
            if (bdayCell.getCellType() == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(bdayCell)) {
                    java.util.Date date = bdayCell.getDateCellValue();
                    java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("MM/dd/yyyy");
                    birthday = sdf.format(date);
                } else {
                    // Serial number - convert manually
                    long serial = (long) bdayCell.getNumericCellValue();
                    java.util.Calendar cal = java.util.Calendar.getInstance();
                    cal.set(1899, java.util.Calendar.DECEMBER, 30);
                    cal.add(java.util.Calendar.DAY_OF_YEAR, (int) serial);
                    java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("MM/dd/yyyy");
                    birthday = sdf.format(cal.getTime());
                }
            }

            double basicSalary = row.getCell(13).getNumericCellValue();
            double hourlyRate = row.getCell(18).getNumericCellValue();

            empNos.add(empNo);
            empNames.add(firstName + " " + lastName);
            empBirthdays.add(birthday);
            basicSalaries.add(basicSalary);
            hourlyRates.add(hourlyRate);
        }

        // --- Sheet 2: Attendance Record ---
        Sheet attSheet = workbook.getSheet("Attendance Record");
        for (int i = 1; i <= attSheet.getLastRowNum(); i++) {
            Row row = attSheet.getRow(i);
            if (row == null) continue;

            int empNo = (int) row.getCell(0).getNumericCellValue();

            // Date (column 3) - get month and day
            Cell dateCell = row.getCell(3);
            int month = 0;
            int day = 0;
            if (dateCell.getCellType() == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(dateCell)) {
                    java.util.Date date = dateCell.getDateCellValue();
                    java.util.Calendar cal = java.util.Calendar.getInstance();
                    cal.setTime(date);
                    month = cal.get(java.util.Calendar.MONTH) + 1;
                    day = cal.get(java.util.Calendar.DAY_OF_MONTH);
                } else {
                    long serial = (long) dateCell.getNumericCellValue();
                    java.util.Calendar cal = java.util.Calendar.getInstance();
                    cal.set(1899, java.util.Calendar.DECEMBER, 30);
                    cal.add(java.util.Calendar.DAY_OF_YEAR, (int) serial);
                    month = cal.get(java.util.Calendar.MONTH) + 1;
                    day = cal.get(java.util.Calendar.DAY_OF_MONTH);
                }
            }

            // Skip invalid months
            if (month < 1 || month > 12) continue;

            // Log In and Log Out (columns 4 and 5) - stored as fraction of day
            double loginFraction = row.getCell(4).getNumericCellValue();
            double logoutFraction = row.getCell(5).getNumericCellValue();

            // Convert fractions to hours
            double loginHours = loginFraction * 24.0;
            double logoutHours = logoutFraction * 24.0;

            // Grace period: if login is at or before 8:10 AM, treat as 8:00 AM
            if (loginHours <= 8.0 + (10.0 / 60.0)) {
                loginHours = 8.0;
            }

            // Cap logout at 5:00 PM (17:00) - do not include extra hours
            if (logoutHours > 17.0) {
                logoutHours = 17.0;
            }

            // Hours worked = adjusted time minus 1 hour lunch break
            double hoursWorked = logoutHours - loginHours - 1.0;
            if (hoursWorked < 0) hoursWorked = 0;

            // Store: {empNo, month, day, hoursWorked}
            attendanceList.add(new double[]{empNo, month, day, hoursWorked});
        }

        workbook.close();
        fis.close();
    }

    // ===================== HELPER METHODS =====================

    /**
     * Find the index of an employee by employee number.
     * Returns -1 if not found.
     */
    static int findEmployee(int empNo) {
        for (int i = 0; i < empNos.size(); i++) {
            if (empNos.get(i) == empNo) return i;
        }
        return -1;
    }

    /**
     * Compute total hours worked for a given employee in a specific month and day range (inclusive).
     */
    static double computeTotalHours(int empNo, int month, int startDay, int endDay) {
        double total = 0;
        for (double[] record : attendanceList) {
            if ((int) record[0] == empNo && (int) record[1] == month
                    && record[2] >= startDay && record[2] <= endDay) {
                total += record[3];
            }
        }
        return total;
    }

    /**
     * SSS Contribution (Employee Share) based on 2024 table.
     * Monthly Salary Credit (MSC) ranges from 4,000 to 30,000.
     * Employee contribution = 4.5% of MSC.
     */
    static double computeSSS(double monthlyBasic) {
        double msc;
        if (monthlyBasic < 4250) {
            msc = 4000;
        } else if (monthlyBasic >= 29750) {
            msc = 30000;
        } else {
            msc = Math.round(monthlyBasic / 500.0) * 500;
        }
        return msc * 0.045;
    }

    /**
     * PhilHealth Contribution (Employee Share) based on 2024 rates.
     * Premium = 5% of monthly basic salary.
     * Employee share = 50% of premium.
     * Premium floor: P500, Salary ceiling: P100,000.
     */
    static double computePhilHealth(double monthlyBasic) {
        double premium = monthlyBasic * 0.05;
        if (premium < 500) premium = 500;
        if (premium > 5000) premium = 5000;
        return premium / 2.0;
    }

    /**
     * Pag-IBIG Contribution (Employee Share).
     * 1% if salary <= 1,500; 2% if salary > 1,500.
     * Maximum employee contribution: P100.
     */
    static double computePagIBIG(double monthlyBasic) {
        double contribution;
        if (monthlyBasic <= 1500) {
            contribution = monthlyBasic * 0.01;
        } else {
            contribution = monthlyBasic * 0.02;
        }
        return Math.min(contribution, 100.0);
    }

    /**
     * Withholding Tax (Monthly) based on TRAIN Law (2023-2024).
     * Taxable income = monthly basic salary - SSS - PhilHealth - Pag-IBIG.
     */
    static double computeTax(double taxableIncome) {
        if (taxableIncome <= 20833) {
            return 0;
        } else if (taxableIncome <= 33333) {
            return (taxableIncome - 20833) * 0.15;
        } else if (taxableIncome <= 66667) {
            return 1875.00 + (taxableIncome - 33333) * 0.20;
        } else if (taxableIncome <= 166667) {
            return 8541.80 + (taxableIncome - 66667) * 0.25;
        } else if (taxableIncome <= 666667) {
            return 33541.80 + (taxableIncome - 166667) * 0.30;
        } else {
            return 183541.80 + (taxableIncome - 666667) * 0.35;
        }
    }

    // Month names and last day of each month (index 1-12)
    static String[] monthNames = {"", "January", "February", "March", "April", "May",
        "June", "July", "August", "September", "October", "November", "December"};
    static int[] lastDays = {0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};

    /**
     * Get the list of months that have attendance data for a given employee.
     */
    static ArrayList<Integer> getMonthsWithData(int empNo) {
        ArrayList<Integer> months = new ArrayList<>();
        for (double[] record : attendanceList) {
            if ((int) record[0] == empNo) {
                int m = (int) record[1];
                if (!months.contains(m)) {
                    months.add(m);
                }
            }
        }
        java.util.Collections.sort(months);
        return months;
    }

    /**
     * Display payroll information for one employee (all months with data).
     */
    static void displayPayroll(int index) {
        int empNo = empNos.get(index);
        String name = empNames.get(index);
        String birthday = empBirthdays.get(index);
        double hourlyRate = hourlyRates.get(index);

        ArrayList<Integer> months = getMonthsWithData(empNo);
        if (months.isEmpty()) {
            System.out.println("No attendance data found for this employee.");
            return;
        }

        for (int month : months) {
            int lastDay = lastDays[month];

            // --- Cutoff 1: 1st to 15th ---
            double hours1 = computeTotalHours(empNo, month, 1, 15);
            double gross1 = hours1 * hourlyRate;

            // --- Cutoff 2: 16th to end of month ---
            double hours2 = computeTotalHours(empNo, month, 16, lastDay);
            double gross2 = hours2 * hourlyRate;

            // Combined gross for deduction computation
            double monthlyGross = gross1 + gross2;

            // Cutoff 1 display
            System.out.println("================================================");
            System.out.printf("Cutoff Date: %s 1 to %s 15%n", monthNames[month], monthNames[month]);
            System.out.println("================================================");
            System.out.printf("Employee #:          %d%n", empNo);
            System.out.printf("Employee Name:       %s%n", name);
            System.out.printf("Birthday:            %s%n", birthday);
            System.out.printf("Total Hours Worked:  %f%n", hours1);
            System.out.printf("Gross Salary:        PHP %,f%n", gross1);
            System.out.printf("Net Salary:          PHP %,f%n", gross1);
            System.out.println();

            // Deductions based on combined cutoff gross
            double sss = computeSSS(monthlyGross);
            double philHealth = computePhilHealth(monthlyGross);
            double pagIbig = computePagIBIG(monthlyGross);
            double taxableIncome = monthlyGross - sss - philHealth - pagIbig;
            double tax = computeTax(taxableIncome);
            double totalDeductions = sss + philHealth + pagIbig + tax;
            double net2 = gross2 - totalDeductions;

            // Cutoff 2 display
            System.out.println("================================================");
            System.out.printf("Cutoff Date: %s 16 to %s %d%n", monthNames[month], monthNames[month], lastDay);
            System.out.println("================================================");
            System.out.printf("Total Hours Worked:  %f%n", hours2);
            System.out.printf("Gross Salary:        PHP %,f%n", gross2);
            System.out.println();
            System.out.println("    Deductions:");
            System.out.printf("    - SSS:               PHP %,f%n", sss);
            System.out.printf("    - PhilHealth:        PHP %,f%n", philHealth);
            System.out.printf("    - Pag-IBIG:          PHP %,f%n", pagIbig);
            System.out.printf("    - Tax:               PHP %,f%n", tax);
            System.out.println();
            System.out.printf("Total Deductions:    PHP %,f%n", totalDeductions);
            System.out.printf("Net Salary:          PHP %,f%n", net2);
            System.out.println("================================================");
            System.out.println();
        }
    }

    // ===================== MAIN PROGRAM =====================

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        // --- LOAD EXCEL FILE ---
        String excelFile = "Copy of MotorPH_Employee Data.xlsx";
        System.out.println("Loading employee data from: " + excelFile);
        try {
            loadExcelData(excelFile);
            System.out.println("Successfully loaded " + empNos.size() + " employees and " + attendanceList.size() + " attendance records.");
        } catch (IOException e) {
            System.out.println("Error: Could not read the Excel file '" + excelFile + "'.");
            System.out.println("Make sure the file is in the same folder as this program.");
            System.out.println("Details: " + e.getMessage());
            scanner.close();
            return;
        }

        // --- LOGIN ---
        System.out.println();
        System.out.println("========================================");
        System.out.println("     MotorPH Payroll System");
        System.out.println("========================================");
        System.out.print("Enter username: ");
        String username = scanner.nextLine().trim();
        System.out.print("Enter password: ");
        String password = scanner.nextLine().trim();

        if (!(username.equals("employee") || username.equals("payroll_staff")) || !password.equals("12345")) {
            System.out.println("Incorrect username and/or password.");
            scanner.close();
            return;
        }

        // --- EMPLOYEE LOGIN ---
        if (username.equals("employee")) {
            boolean running = true;
            while (running) {
                System.out.println();
                System.out.println("========================================");
                System.out.println("          Employee Portal");
                System.out.println("========================================");
                System.out.println("1. Enter your employee number");
                System.out.println("2. Exit the program");
                System.out.print("Choose an option: ");
                String choice = scanner.nextLine().trim();

                switch (choice) {
                    case "1":
                        System.out.print("Enter employee number: ");
                        String empInput = scanner.nextLine().trim();
                        try {
                            int empNo = Integer.parseInt(empInput);
                            int index = findEmployee(empNo);
                            if (index == -1) {
                                System.out.println("Employee number does not exist.");
                            } else {
                                System.out.println();
                                System.out.println("----------------------------------------");
                                System.out.printf("Employee Number: %d%n", empNos.get(index));
                                System.out.printf("Employee Name:   %s%n", empNames.get(index));
                                System.out.printf("Birthday:        %s%n", empBirthdays.get(index));
                                System.out.println("----------------------------------------");
                            }
                        } catch (NumberFormatException e) {
                            System.out.println("Employee number does not exist.");
                        }
                        break;
                    case "2":
                        running = false;
                        System.out.println("Exiting the program. Goodbye!");
                        break;
                    default:
                        System.out.println("Invalid option. Please try again.");
                }
            }
        }

        // --- PAYROLL STAFF LOGIN ---
        if (username.equals("payroll_staff")) {
            boolean running = true;
            while (running) {
                System.out.println();
                System.out.println("========================================");
                System.out.println("       Payroll Staff Portal");
                System.out.println("========================================");
                System.out.println("1. Process Payroll");
                System.out.println("2. Exit the program");
                System.out.print("Choose an option: ");
                String choice = scanner.nextLine().trim();

                switch (choice) {
                    case "1":
                        boolean processing = true;
                        while (processing) {
                            System.out.println();
                            System.out.println("----------------------------------------");
                            System.out.println("        Process Payroll");
                            System.out.println("----------------------------------------");
                            System.out.println("1. One employee");
                            System.out.println("2. All employees");
                            System.out.println("3. Exit the program");
                            System.out.print("Choose an option: ");
                            String subChoice = scanner.nextLine().trim();

                            switch (subChoice) {
                                case "1":
                                    boolean lookingUp = true;
                                    while (lookingUp) {
                                        System.out.println();
                                        System.out.print("Enter employee number: ");
                                        String empInput = scanner.nextLine().trim();
                                        try {
                                            int empNo = Integer.parseInt(empInput);
                                            int index = findEmployee(empNo);
                                            if (index == -1) {
                                                System.out.println("Employee number does not exist.");
                                            } else {
                                                System.out.println();
                                                displayPayroll(index);
                                            }
                                        } catch (NumberFormatException e) {
                                            System.out.println("Employee number does not exist.");
                                        }

                                        System.out.println();
                                        System.out.println("a. Enter another employee number");
                                        System.out.println("b. Exit the program");
                                        System.out.print("Choose an option: ");
                                        String empChoice = scanner.nextLine().trim().toLowerCase();

                                        if (empChoice.equals("b")) {
                                            lookingUp = false;
                                            processing = false;
                                            running = false;
                                            System.out.println("Exiting the program. Goodbye!");
                                        } else if (!empChoice.equals("a")) {
                                            System.out.println("Invalid option. Please try again.");
                                        }
                                    }
                                    break;
                                case "2":
                                    System.out.println();
                                    for (int i = 0; i < empNos.size(); i++) {
                                        displayPayroll(i);
                                    }
                                    break;
                                case "3":
                                    processing = false;
                                    running = false;
                                    System.out.println("Exiting the program. Goodbye!");
                                    break;
                                default:
                                    System.out.println("Invalid option. Please try again.");
                            }
                        }
                        break;
                    case "2":
                        running = false;
                        System.out.println("Exiting the program. Goodbye!");
                        break;
                    default:
                        System.out.println("Invalid option. Please try again.");
                }
            }
        }

        scanner.close();
    }
}
