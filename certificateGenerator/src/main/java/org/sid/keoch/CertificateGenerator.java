package org.sid.keoch;

//  String excelFilePath = "D:\\Documents\\Diploma\\TEMPLATE\\student_template.xlsx"; // Path to Excel file
//        String wordTemplatePath = "D:\\Documents\\Diploma\\TEMPLATE\\certificate_template.docx"; // Path to Word template
//        String outputFolder = "D:\\Documents\\Diploma\\TEMPLATE\\certificates_try\\"; // Folder to save generated certificates

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CertificateGenerator {
    private static final Logger logger = LoggerFactory.getLogger(CertificateGenerator.class);

    public static void main(String[] args) {
        String excelFilePath = "D:\\Documents\\Diploma\\TEMPLATE\\student_template.xlsx"; // Path to Excel file
        String wordTemplatePath = "D:\\Documents\\Diploma\\TEMPLATE\\certificate_template.docx"; // Path to Word template
        String outputFolder = "D:\\Documents\\Diploma\\TEMPLATE\\certificates_try\\"; // Folder to save generated certificates

        try {
            // Create output folder if it doesn't exist
            File folder = new File(outputFolder);
            if (!folder.exists()) {
                folder.mkdirs();
            }

            // Read data from Excel file
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through rows in the Excel sheet
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next(); // Skip header row

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Read data from each column using a helper method
                String studentName = getCellValueAsString(row.getCell(0));
                String registrationNo = getCellValueAsString(row.getCell(1));
                String collegeName = getCellValueAsString(row.getCell(2));
                String internshipDomain = getCellValueAsString(row.getCell(3));
                String issuanceDate = getCellValueAsString(row.getCell(4));
                String fromDate = getCellValueAsString(row.getCell(5));
                String toDate = getCellValueAsString(row.getCell(6));

                // Log the values read from Excel
                logger.info("Student Name: {}", studentName);
                logger.info("Registration No: {}", registrationNo);
                logger.info("College Name: {}", collegeName);
                logger.info("Internship Domain: {}", internshipDomain);
                logger.info("Issuance Date: {}", issuanceDate);
                logger.info("From Date: {}", fromDate);
                logger.info("To Date: {}", toDate);

                // Create a map of placeholders and their replacements
                Map<String, String> placeholders = new HashMap<>();
                placeholders.put("{{student_name}}", studentName);
                placeholders.put("{{registration_no}}", registrationNo);
                placeholders.put("{{college_name}}", collegeName);
                placeholders.put("{{Internship_area}}", internshipDomain);
                placeholders.put("{{issuance_date}}", issuanceDate);
                placeholders.put("{{from_date}}", fromDate);
                placeholders.put("{{to_date}}", toDate);

                // Generate certificate for the student
                generateCertificate(wordTemplatePath, outputFolder + studentName + "_certificate.docx", placeholders);
            }

            workbook.close();
            System.out.println("Certificates generated successfully!");

        } catch (IOException e ) {
            logger.error("An error occurred: ", e);
        }
    }

    private static void generateCertificate(String templatePath, String outputPath, Map<String, String> placeholders) throws IOException {
        FileInputStream templateFile = new FileInputStream(new File(templatePath));
        XWPFDocument document = new XWPFDocument(templateFile);

        // Replace placeholders in the document
        replacePlaceholders(document, placeholders);

        // Save the modified document
        FileOutputStream outFile = new FileOutputStream(new File(outputPath));
        document.write(outFile);
        outFile.close();
        document.close();
    }

    private static void replacePlaceholders(XWPFDocument document, Map<String, String> placeholders) {
        // Iterate through paragraphs
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            replacePlaceholdersInParagraph(paragraph, placeholders);
        }

        // Iterate through tables
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
        try {
            String paragraphText = paragraph.getText();
            logger.debug("Original paragraph text: {}", paragraphText);

            // Iterate through each placeholder
            for (Map.Entry<String, String> entry : placeholders.entrySet()) {
                String placeholder = entry.getKey();
                String value = entry.getValue();

                // Check if the paragraph text contains the placeholder
                if (paragraphText.contains(placeholder)) {
                    logger.info("Replacing placeholder {} with value {}", placeholder, value);

                    // Get the runs of the paragraph
                    List<XWPFRun> runs = paragraph.getRuns();
                    logger.debug("Number of runs in the paragraph: {}", runs.size());

                    // Clear the existing runs
                    for (int i = runs.size() - 1; i >= 0; i--) {
                        paragraph.removeRun(i);
                    }

                    // Add the new run with replaced text
                    XWPFRun run = paragraph.createRun();
                    run.setText(paragraphText.replace(placeholder, value));
                    run.setBold(true);
                    logger.debug("Replaced text: {}", paragraphText.replace(placeholder, value));
                } else {
                    logger.debug("Placeholder {} not found in the paragraph", placeholder);
                }
            }
        } catch (Exception e) {
            logger.error("Error processing paragraph: ", e);
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return formatDate(cell.getDateCellValue());
                } else {
                    return String.valueOf(cell.getNumericCellValue()).trim();
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).trim();
            case FORMULA:
                return cell.getCellFormula().trim();
            default:
                return "";
        }
    }

    private static String formatDate(java.util.Date date) {
        if (date == null) {
            return "";
        }
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd MMMM yyyy");
        return date.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate().format(formatter);
    }
}
