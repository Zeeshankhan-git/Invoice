package org.TSMM;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;

public class TMSS_Invoice_Student_Generator {

    public static void main(String[] args) {
        String excelFilePath = "Students.xlsx"; // Update with correct file path
        java.util.List<String[]> students = readStudentsFromExcel(excelFilePath);

        if (students.isEmpty()) {
            System.out.println("No students found in the Excel file.");
            return;
        }

        for (String[] student : students) {
            String studentName = student[1].replace(" ", "_"); // Replace spaces with underscores
            String pdfFilePath = "Invoice_" + studentName + ".pdf";
            generatePdfInvoice(pdfFilePath, student);
        }
    }

    // Read Student Data from Excel
    private static java.util.List<String[]> readStudentsFromExcel(String excelFilePath) {
        java.util.List<String[]> students = new java.util.ArrayList<>();

        try (FileInputStream file = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0);
            boolean isHeader = true; // Flag to skip the first row

            for (Row row : sheet) {
                if (isHeader) {
                    isHeader = false; // Skip the header row
                    continue;
                }

                Cell nameCell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                Cell idCell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                Cell phoneCell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                String name = nameCell.toString().trim();
                String studentId = idCell.toString().trim();
                String phone = phoneCell.toString().trim();

                if (!name.isEmpty() && !studentId.isEmpty()) {
                    students.add(new String[]{name, studentId, phone});
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return students;
    }


    // Generate PDF Invoice
    private static void generatePdfInvoice(String filePath, String[] student) {
        Document document = new Document();
        try {
            PdfWriter.getInstance(document, new FileOutputStream(filePath));
            document.open();

            // Add Company Logo
            // Add Company Logo
            String logoPath = "src/main/resources/TMSS LOGO.jpg"; // Adjust the path as needed
            Image logo = Image.getInstance(logoPath);
            logo.scaleToFit(200, 50);
            logo.setAlignment(Element.ALIGN_CENTER);
            document.add(logo);

            // Fonts
            com.itextpdf.text.Font titleFont = new com.itextpdf.text.Font(com.itextpdf.text.Font.FontFamily.HELVETICA, 16, com.itextpdf.text.Font.BOLD);
            com.itextpdf.text.Font normalFont = new com.itextpdf.text.Font(com.itextpdf.text.Font.FontFamily.HELVETICA, 12);
            com.itextpdf.text.Font boldFont = new com.itextpdf.text.Font(com.itextpdf.text.Font.FontFamily.HELVETICA, 12, com.itextpdf.text.Font.BOLD);

            // Invoice Header
            document.add(new Paragraph("INVOICE", titleFont));
            document.add(Chunk.NEWLINE);
            document.add(new Paragraph("Invoice No.: TMSS/2024-2025/DPSK/INV/" + student[2], normalFont));
            document.add(new Paragraph("Invoice Date: 30-Nov-2024", normalFont));
            document.add(Chunk.NEWLINE);

            // Company Info
            document.add(new Paragraph("From:", boldFont));
            document.add(new Paragraph("TechnoMedia Software Solutions Pvt. Ltd.", normalFont));
            document.add(new Paragraph("CV Ramannagar, Bangalore, Karnataka - 560075", normalFont));
            document.add(new Paragraph("Email: info@technomediasoft.com", normalFont));
            document.add(Chunk.NEWLINE);

            // Bill To
            document.add(new Paragraph("Bill To: " + student[0],boldFont));

            document.add(Chunk.NEWLINE);

            // Invoice Items Table
            PdfPTable table = new PdfPTable(3);
            table.setWidthPercentage(100);
            table.addCell(new PdfPCell(new Phrase("Sl. No.", boldFont)));
            table.addCell(new PdfPCell(new Phrase("Description", boldFont)));
            table.addCell(new PdfPCell(new Phrase("Amount (Rs.)", boldFont)));

            int baseAmount = 1000; // Example amount
            table.addCell("1");
            table.addCell("RouteAlert charges for Nov-2024\nNumber Of Students: 748");
            table.addCell(String.valueOf(baseAmount));
            document.add(table);
            document.add(Chunk.NEWLINE);

            // Taxes Calculation
            double discount = 0.00;
            double subtotal = baseAmount - discount;
            double igst = subtotal * 0.18;
            double totalAmount = subtotal + igst;

            // Tax Breakdown
            document.add(new Paragraph("Other Charges:", boldFont));
            document.add(new Paragraph("Discount: Rs. " + discount, normalFont));
            document.add(new Paragraph("Subtotal: Rs. " + subtotal, normalFont));
            document.add(new Paragraph("IGST @ 18%: Rs. " + igst, normalFont));
            document.add(new Paragraph("Total: Rupees " + totalAmount + " only", boldFont));
            document.add(Chunk.NEWLINE);

            // Generate & Add QR Code
            String qrFileName = Paths.get(System.getProperty("user.dir"), "QR_" + student[2] + ".png").toString();
            generateQRCode(student[2], qrFileName);
            try {
                Image qrImage = Image.getInstance(qrFileName);
                qrImage.scaleToFit(100, 100);
                qrImage.setAlignment(Element.ALIGN_CENTER);
                document.add(qrImage);
            } catch (IOException e) {
                System.out.println(" QR Code not found: " + qrFileName);
            }

            document.add(Chunk.NEWLINE);

            // Bank Details
            document.add(new Paragraph("Bank Details:", boldFont));
            document.add(new Paragraph("TechnoMedia Software Solutions Pvt. Ltd.", normalFont));
            document.add(new Paragraph("Bank: Indian Overseas Bank", normalFont));
            document.add(new Paragraph("Branch: Bangalore-560075", normalFont));
            document.add(Chunk.NEWLINE);

            // Authorized Signatory
            document.add(new Paragraph("For TechnoMedia Software Solutions Pvt. Ltd.", boldFont));
            document.add(new Paragraph("Authorized Signatory", titleFont));

            document.close();
            System.out.println(" Invoice Created: " + filePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Generate QR Code
    private static void generateQRCode(String studentId, String fileName) {
        int width = 200;
        int height = 200;
        try {
            BitMatrix bitMatrix = new MultiFormatWriter().encode(
                    "Invoice for Student ID: " + studentId,
                    BarcodeFormat.QR_CODE, width, height
            );

            Path path = Paths.get(fileName);
            MatrixToImageWriter.writeToPath(bitMatrix, "PNG", path);
            System.out.println(" QR Code Generated: " + fileName);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
