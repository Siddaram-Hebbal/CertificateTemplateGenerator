package org.sid.keoch;

//  String excelFilePath = "D:\\Documents\\Diploma\\TEMPLATE\\student_template.xlsx"; // Path to Excel file
//        String wordTemplatePath = "D:\\Documents\\Diploma\\TEMPLATE\\certificate_template.docx"; // Path to Word template
//        String outputFolder = "D:\\Documents\\Diploma\\TEMPLATE\\certificates_try\\"; // Folder to save generated certificates


import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class CertificateGeneratorPDF {
    public static void main(String[] args) {
        String excelFilePath = "D:\\Documents\\Diploma\\TEMPLATE\\student_template.xlsx";
        String wordTemplatePath = "D:\\Documents\\Diploma\\TEMPLATE\\certificate_template.docx";
        String outputFolder = "D:\\Documents\\Diploma\\TEMPLATE\\certificates\\";

        try {
            Files.createDirectories(Paths.get(outputFolder));

            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next(); // Skip header row

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Read student details
                String studentName = getCellValueAsString(row.getCell(0));
                String registrationNo = getCellValueAsString(row.getCell(1));
                String collegeName = getCellValueAsString(row.getCell(2));
                String internshipDomain = getCellValueAsString(row.getCell(3));
                String issuanceDate = getCellValueAsString(row.getCell(4));
                String fromDate = getCellValueAsString(row.getCell(5));
                String toDate = getCellValueAsString(row.getCell(6));

                // Map placeholders to values
                Map<String, String> placeholders = new HashMap<>();
                placeholders.put("{{student_name}}", studentName);
                placeholders.put("{{registration_no}}", registrationNo);
                placeholders.put("{{college_name}}", collegeName);
                placeholders.put("{{Internship_area}}", internshipDomain);
                placeholders.put("{{issuance_date}}", issuanceDate);
                placeholders.put("{{from_date}}", fromDate);
                placeholders.put("{{to_date}}", toDate);

                // Generate Word Certificate
                String docxPath = outputFolder + studentName + "_certificate.docx";
                generateCertificate(wordTemplatePath, docxPath, placeholders);

                // Convert Word to PDF
                String pdfPath = outputFolder + studentName + "_certificate.pdf";
                convertDocxToPdf(docxPath, pdfPath);
            }

            workbook.close();
            System.out.println("Certificates generated successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void generateCertificate(String templatePath, String outputPath, Map<String, String> placeholders) throws IOException {
        FileInputStream templateFile = new FileInputStream(new File(templatePath));
        XWPFDocument document = new XWPFDocument(templateFile);

        replacePlaceholders(document, placeholders);

        FileOutputStream outFile = new FileOutputStream(new File(outputPath));
        document.write(outFile);
        outFile.close();
        document.close();
    }

    private static void replacePlaceholders(XWPFDocument document, Map<String, String> placeholders) {
        // Replace in paragraphs
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            replacePlaceholdersInParagraph(paragraph, placeholders);
        }

        // Replace in tables
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        replacePlaceholdersInParagraph(paragraph, placeholders);
                    }
                }
            }
        }
    }

    private static void replacePlaceholdersInParagraph(XWPFParagraph paragraph, Map<String, String> placeholders) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null) return;

        StringBuilder text = new StringBuilder();
        for (XWPFRun run : runs) {
            text.append(run.text());
        }

        String replacedText = text.toString();
        boolean replaced = false;

        for (Map.Entry<String, String> entry : placeholders.entrySet()) {
            if (replacedText.contains(entry.getKey())) {
                replacedText = replacedText.replace(entry.getKey(), entry.getValue());
                replaced = true;
            }
        }

        if (replaced) {
            while (paragraph.getRuns().size() > 0) {
                paragraph.removeRun(0);
            }
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(replacedText);
            newRun.setBold(true); // Ensuring bold format is applied
        }
    }

    private static void convertDocxToPdf(String docxPath, String pdfPath) {
        try {
            Document document = new Document();
            document.loadFromFile(docxPath);
            document.saveToFile(pdfPath, FileFormat.PDF);
            System.out.println("Converted " + docxPath + " to " + pdfPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((long) cell.getNumericCellValue()).trim();
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).trim();
            default:
                return "";
        }
    }
}
