package com.demo.ExcelDemo;
import org.apache.log4j.BasicConfigurator;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class EmployeeAnalyzer {

    public static void main(String[] args) {
    	BasicConfigurator.configure();
        String inputFile = "C:\\assignment.xlsx";

        try {
            FileInputStream file = new FileInputStream(new File(inputFile));
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0); 

            
            analyzeConsecutiveDays(sheet);
            analyzeShiftGaps(sheet);
            analyzeLongShifts(sheet);

            workbook.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void analyzeConsecutiveDays(Sheet sheet) {
        System.out.println("Employees who worked for 7 consecutive days:");
        Map<String, Integer> consecutiveDaysMap = new HashMap<>();

       
        Iterator<Row> iterator = sheet.iterator();
        if (iterator.hasNext()) {
            iterator.next(); // Skip header row
        }

        while (iterator.hasNext()) {
            Row row = iterator.next();
            Cell employeeNameCell = row.getCell(7); // Assuming Employee Name is in column 8 (index 7)
            if (employeeNameCell == null || employeeNameCell.getCellType() == CellType.BLANK) {
                continue; // Skip empty rows
            }

            String employeeName = null;
            if (employeeNameCell.getCellType() == CellType.STRING) {
                employeeName = employeeNameCell.getStringCellValue();
            } else if (employeeNameCell.getCellType() == CellType.NUMERIC) {
                employeeName = String.valueOf((int) employeeNameCell.getNumericCellValue());
            }

            Cell timeInCell = row.getCell(2); // Assuming Time In is in column 3 (index 2)
            if (timeInCell == null || timeInCell.getCellType() == CellType.BLANK) {
                continue; // Skip rows with missing Time In data
            }

            String timeInText = null;
            if (timeInCell.getCellType() == CellType.STRING) {
                timeInText = timeInCell.getStringCellValue();
            } else if (timeInCell.getCellType() == CellType.NUMERIC) {
                timeInText = String.valueOf((int) timeInCell.getNumericCellValue());
            }

            try {
                LocalDateTime timeIn = LocalDateTime.parse(timeInText, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
                consecutiveDaysMap.put(employeeName, consecutiveDaysMap.getOrDefault(employeeName, 0) + 1);
            } catch (DateTimeParseException e) {
                System.out.println("Skipping row with invalid time format: " + timeInText);
            }
        }

        for (Map.Entry<String, Integer> entry : consecutiveDaysMap.entrySet()) {
            if (entry.getValue() >= 7) {
                System.out.println("Employee: " + entry.getKey() + ", Consecutive Days: " + entry.getValue());
            }
        }
    }


    private static void analyzeShiftGaps(Sheet sheet) {
        System.out.println("Employees with less than 10 hours between shifts but greater than 1 hour:");
        Map<String, LocalDateTime> lastShiftEnd = new HashMap<>();

        // Skip the first row assuming it's a header row
        Iterator<Row> iterator = sheet.iterator();
        if (iterator.hasNext()) {
            iterator.next(); // Skip header row
        }

        while (iterator.hasNext()) {
            Row row = iterator.next();
            Cell employeeNameCell = row.getCell(7); // Assuming Employee Name is in column 8 (index 7)
            if (employeeNameCell == null || employeeNameCell.getCellType() == CellType.BLANK) {
                continue; // Skip empty rows
            }

            String employeeName = null;
            if (employeeNameCell.getCellType() == CellType.STRING) {
                employeeName = employeeNameCell.getStringCellValue();
            } else if (employeeNameCell.getCellType() == CellType.NUMERIC) {
                employeeName = String.valueOf((int) employeeNameCell.getNumericCellValue());
            }

            Cell timeInCell = row.getCell(2); // Assuming Time In is in column 3 (index 2)
            if (timeInCell == null || timeInCell.getCellType() == CellType.BLANK) {
                continue; // Skip rows with missing Time In data
            }

            String timeInText = null;
            if (timeInCell.getCellType() == CellType.STRING) {
                timeInText = timeInCell.getStringCellValue();
            } else if (timeInCell.getCellType() == CellType.NUMERIC) {
                timeInText = String.valueOf((int) timeInCell.getNumericCellValue());
            }

            try {
                LocalDateTime timeIn = LocalDateTime.parse(timeInText, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));

                if (lastShiftEnd.containsKey(employeeName)) {
                    LocalDateTime lastEnd = lastShiftEnd.get(employeeName);
                    Duration duration = Duration.between(lastEnd, timeIn);
                    if (duration.toHours() > 1 && duration.toHours() < 10) {
                        System.out.println("Employee: " + employeeName + ", Gap between shifts: " + duration.toHours() + " hours");
                    }
                }

                lastShiftEnd.put(employeeName, timeIn);
            } catch (DateTimeParseException e) {
                System.out.println("Skipping row with invalid time format: " + timeInText);
            }
        }
    }


    private static void analyzeLongShifts(Sheet sheet) {
        System.out.println("Employees who worked for more than 14 hours in a single shift:");

        // Skip the first row assuming it's a header row
        Iterator<Row> iterator = sheet.iterator();
        if (iterator.hasNext()) {
            iterator.next(); 
        }

        while (iterator.hasNext()) {
            Row row = iterator.next();
            Cell employeeNameCell = row.getCell(7); // Assuming Employee Name is in column 8 (index 7)
            if (employeeNameCell == null || employeeNameCell.getCellType() == CellType.BLANK) {
                continue; // Skip empty rows
            }

            String employeeName = null;
            if (employeeNameCell.getCellType() == CellType.STRING) {
                employeeName = employeeNameCell.getStringCellValue();
            } else if (employeeNameCell.getCellType() == CellType.NUMERIC) {
                employeeName = String.valueOf((int) employeeNameCell.getNumericCellValue());
            }

            Cell timeInCell = row.getCell(2); 
            Cell timeOutCell = row.getCell(3); 

            if (timeInCell == null || timeInCell.getCellType() == CellType.BLANK ||
                timeOutCell == null || timeOutCell.getCellType() == CellType.BLANK) {
                continue; 
            }

            String timeInText = null;
            String timeOutText = null;
            if (timeInCell.getCellType() == CellType.STRING && timeOutCell.getCellType() == CellType.STRING) {
                timeInText = timeInCell.getStringCellValue();
                timeOutText = timeOutCell.getStringCellValue();
            } else {
                continue; 
            }

            try {
                LocalDateTime timeIn = LocalDateTime.parse(timeInText, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
                LocalDateTime timeOut = LocalDateTime.parse(timeOutText, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));

                Duration duration = Duration.between(timeIn, timeOut);
                if (duration.toHours() > 14) {
                    System.out.println("Employee: " + employeeName + ", Shift duration: " + duration.toHours() + " hours");
                }
            } catch (DateTimeParseException e) {
                System.out.println("Skipping row with invalid time format: " + timeInText + " - " + timeOutText);
            }
        }
    }}
