package org.TSMM;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class TMSSInvoiceExcel {
    public static void main(String[] args) {
        String studentFilePath = "/mnt/data/Students.xlsx"; // Path of uploaded file
        String invoiceFilePath = "TMSS_Invoice.xlsx";

        // Read student data from Students.xlsx
        List<String[]> students = readStudentsFromExcel(studentFilePath);
        Random random = new Random();

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Invoice");

            // Create Title Style
            CellStyle titleStyle = workbook.createCellStyle();
            Font titleFont = workbook.createFont();
            titleFont.setBold(true);
            titleFont.setFontHeightInPoints((short) 14);
            titleStyle.setFont(titleFont);

            // Create Header Style
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);

            int rowNum = 0;
            Row row = sheet.createRow(rowNum++);
            Cell cell = row.createCell(1);
            cell.setCellValue("INVOICE");
            cell.setCellStyle(titleStyle);

            rowNum++;
            sheet.createRow(rowNum++).createCell(0).setCellValue("Invoice No.: TMSS/2024-2025/DPSK/INV/15");
            sheet.createRow(rowNum++).createCell(0).setCellValue("Invoice Date: 30-Nov-2024");

            rowNum++;
            sheet.createRow(rowNum++).createCell(0).setCellValue("From: TechnoMedia Software Solutions Pvt. Ltd.");
            sheet.createRow(rowNum++).createCell(0).setCellValue("CV Ramannagar, Bangalore, Karnataka - 560075");
            sheet.createRow(rowNum++).createCell(0).setCellValue("Email: info@technomediasoft.com");
            sheet.createRow(rowNum++).createCell(0).setCellValue("PAN: AACCXXXXXXX");
            sheet.createRow(rowNum++).createCell(0).setCellValue("GSTN No.: 29AACXXXXXXXXX");

            rowNum++;
            sheet.createRow(rowNum++).createCell(0).setCellValue("Bill To: DPS Khanapara, Guwahati");

            rowNum++;
            Row headerRow = sheet.createRow(rowNum++);
            headerRow.createCell(0).setCellValue("Sl. No.");
            headerRow.createCell(1).setCellValue("Student ID");
            headerRow.createCell(2).setCellValue("Student Name");
            headerRow.createCell(3).setCellValue("Amount (Rs.)");

            // Apply header style
            for (int i = 0; i < 4; i++) {
                headerRow.getCell(i).setCellStyle(headerStyle);
            }

            // Insert Student Data with Random Amounts
            int serialNumber = 1;
            for (String[] student : students) {
                Row dataRow = sheet.createRow(rowNum++);
                dataRow.createCell(0).setCellValue(serialNumber++);
                dataRow.createCell(1).setCellValue(student[0]); // Student ID
                dataRow.createCell(2).setCellValue(student[1]); // Student Name
                dataRow.createCell(3).setCellValue(random.nextInt(5000) + 500); // Random amount between 500-5500
            }

            rowNum++;
            sheet.createRow(rowNum++).createCell(0).setCellValue("Other Charges: Discount:");
            sheet.createRow(rowNum++).createCell(0).setCellValue("Subtotal: Rs. 0.00");
            sheet.createRow(rowNum++).createCell(0).setCellValue("CGST @ 0%: Rs. 0.00");
            sheet.createRow(rowNum++).createCell(0).setCellValue("SGST @ 0%: Rs. 0.00");
            sheet.createRow(rowNum++).createCell(0).setCellValue("IGST @ 18%: Rs. 0.00");
            sheet.createRow(rowNum++).createCell(0).setCellValue("Total Tax: Rs. 0.00");
            sheet.createRow(rowNum++).createCell(0).setCellValue("Total: Rupees ZERO only");

            rowNum++;
            sheet.createRow(rowNum++).createCell(0).setCellValue("For TechnoMedia Software Solutions Pvt. Ltd.");
            sheet.createRow(rowNum++).createCell(0).setCellValue("Authorized Signatory");

            // Auto-size columns for better readability
            for (int i = 0; i < 4; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write the output to the file
            try (FileOutputStream fileOut = new FileOutputStream(invoiceFilePath)) {
                workbook.write(fileOut);
            }

            System.out.println("✅ Invoice Excel file created successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Function to read student details from Excel
    private static List<String[]> readStudentsFromExcel(String excelFilePath) {
        List<String[]> students = new ArrayList<>();

        try (FileInputStream file = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header row

                Cell idCell = row.getCell(0);
                Cell nameCell = row.getCell(1);

                if (idCell != null && nameCell != null) {
                    String studentId = idCell.toString().trim();
                    String studentName = nameCell.toString().trim();
                    students.add(new String[]{studentId, studentName});
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return students;
    }
}
