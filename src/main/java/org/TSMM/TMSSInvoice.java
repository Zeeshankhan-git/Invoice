package org.TSMM;


import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import java.io.FileOutputStream;

public class TMSSInvoice {
    public static void main(String[] args) {
        String dest = "TMSS_Invoice_Generated.pdf";
        try {
            Document document = new Document();
            PdfWriter.getInstance(document, new FileOutputStream(dest));
            document.open();

            // adding logo in the invoice
            Image logo = Image.getInstance("TMSS LOGO.jpg");
            logo.scaleToFit(150, 50); // Adjust the logo size
            logo.setAlignment(Element.ALIGN_CENTER);
            document.add(logo);
            document.add(Chunk.NEWLINE);

            // adding the Title
            Font titleFont = new Font(Font.FontFamily.HELVETICA, 16, Font.BOLD);
            Paragraph title = new Paragraph("INVOICE", titleFont);
            title.setAlignment(Element.ALIGN_CENTER);
            document.add(title);
            document.add(Chunk.NEWLINE);

            // Invoice details
            document.add(new Paragraph("Invoice No.: TMSS/2024-2025/DPSK/INV/15"));
            document.add(new Paragraph("Invoice Date: 30-Nov-2024"));
            document.add(Chunk.NEWLINE);

            // From details
            document.add(new Paragraph("From: TechnoMedia Software Solutions Pvt. Ltd."));
            document.add(new Paragraph("CV Ramannagar, Bangalore, Karnataka - 560075"));
            document.add(new Paragraph("Email: info@technomediasoft.com"));
            document.add(new Paragraph("PAN: AACCXXXXXXX"));
            document.add(new Paragraph("GSTN No.: 29AACXXXXXXXXX"));
            document.add(Chunk.NEWLINE);

            // Bill To details
            document.add(new Paragraph("Bill To: DPS Khanapara, Guwahati"));
            document.add(Chunk.NEWLINE);

            // Table for invoice items
            PdfPTable table = new PdfPTable(3);
            table.setWidthPercentage(100);
            table.addCell("Sl. No.");
            table.addCell("Description");
            table.addCell("Amount (Rs.)");
            table.addCell("1");
            table.addCell("RouteAlert charges for the month of Nov-2024\nNumber Of Students: 748");
            table.addCell("0.00");
            document.add(table);
            document.add(Chunk.NEWLINE);

            // Other charges
//            document.add(new Paragraph("Other Charges: Discount:"));
//            document.add(new Paragraph("Subtotal: Rs. 0.00"));
//            document.add(new Paragraph("CGST @ 0%: Rs. 0.00"));
//            document.add(new Paragraph("SGST @ 0%: Rs. 0.00"));
//            document.add(new Paragraph("IGST @ 18%: Rs. 0.00"));
//            document.add(new Paragraph("Total Tax: Rs. 0.00"));
//            document.add(new Paragraph("Total: Rupees ZERO only"));
//            document.add(Chunk.NEWLINE);


            // Define a bold font
            Font boldFont = new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD);

// Add invoice summary with bold text
            document.add(new Paragraph("Other Charges:", boldFont));
            document.add(new Paragraph("Discount:", boldFont));
            document.add(new Paragraph("Subtotal: Rs. 0.00", boldFont));
            document.add(new Paragraph("CGST @ 0%: Rs. 0.00", boldFont));
            document.add(new Paragraph("SGST @ 0%: Rs. 0.00", boldFont));
            document.add(new Paragraph("IGST @ 18%: Rs. 0.00", boldFont));
            document.add(new Paragraph("Total Tax: Rs. 0.00", boldFont));
            document.add(new Paragraph("Total: RUPEES ZERO ONLY", boldFont));


            // Bank Details
            document.add(new Paragraph("Please CREDIT:"));
            document.add(new Paragraph("Account Name: TechnoMedia Software Solutions Pvt. Ltd."));
            document.add(new Paragraph("Account Number: 160002000000446"));
            document.add(new Paragraph("IFS Code: IOBA0001600"));
            document.add(new Paragraph("Bank: Indian Overseas Bank"));
            document.add(new Paragraph("Branch: Vigyana Nagar Branch, Bangalore-560075"));
            document.add(new Paragraph("Bank Phone No: +91-80-22950165/6"));
            document.add(Chunk.NEWLINE);

            // Authorized Signatory
            document.add(new Paragraph("For TechnoMedia Software Solutions Pvt. Ltd."));
            document.add(Chunk.NEWLINE);
            document.add(new Paragraph("Authorized Signatory", titleFont));

            document.close();
            System.out.println("Invoice PDF with logo created successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
