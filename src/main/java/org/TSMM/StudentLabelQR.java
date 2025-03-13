package org.TSMM;


import com.google.zxing.BarcodeFormat;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class StudentLabelQR {
    public static void main(String[] args) {
        String excelFilePath = "TMSSInvoiceExcel.xlsx";  // Ensure this file is in your project directory
        String htmlFileName = "TMSSInvoiceExcel.html";

        // Extract student data from Excel
        List<String[]> students = readStudentsFromExcel(excelFilePath);

        if (students.isEmpty()) {
            System.out.println("No student data found in Excel.");
            return;
        }

        try (FileWriter writer = new FileWriter(htmlFileName)) {
            writer.write("<html><head><title>Student Label Print</title>");
            writer.write("<style>");
            writer.write("body { font-family: Arial, sans-serif; text-align: center; }");
            writer.write(".label { display: inline-block; width: 250px; border: 1px solid #000; padding: 10px; margin: 10px; }");
            writer.write(".qr { width: 100px; height: 100px; }");
            writer.write("</style>");
            writer.write("</head><body>");
            writer.write("<h2>Student Labels with QR Codes</h2>");

            // Generate QR Codes and Create Labels
            for (String[] student : students) {
                String studentId = student[0];
                String studentName = student[1];
                String qrFileName = "QR_" + studentId + ".png";

                // Generate QR Code
                generateQRCode(studentId, qrFileName);

                // Write HTML Content
                writer.write("<div class='label'>");
                writer.write("<p><strong>ID:</strong> " + studentId + "</p>");
                writer.write("<p><strong>Name:</strong> " + studentName + "</p>");
                writer.write("<img class='qr' src='" + qrFileName + "' alt='QR Code'>");
                writer.write("</div>");
            }

            writer.write("</body></html>");
            System.out.println("Student Labels HTML file created successfully!");

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

    // Function to generate QR codes
    private static void generateQRCode(String data, String filePath) {
        try {
            int width = 100;
            int height = 100;
            BitMatrix bitMatrix = new MultiFormatWriter().encode(data, BarcodeFormat.QR_CODE, width, height);
            MatrixToImageWriter.writeToPath(bitMatrix, "PNG", Paths.get(filePath));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
