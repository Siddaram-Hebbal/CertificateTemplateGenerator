package org.sid.keoch;

//  String excelFilePath = "D:\\Documents\\Diploma\\TEMPLATE\\student_template.xlsx"; // Path to Excel file
//        String wordTemplatePath = "D:\\Documents\\Diploma\\TEMPLATE\\certificate_template.docx"; // Path to Word template
//        String outputFolder = "D:\\Documents\\Diploma\\TEMPLATE\\certificates_try\\"; // Folder to save generated certificates
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class CertificateGeneratorTest {
    private static final Logger logger = LoggerFactory.getLogger(CertificateGenerator.class);

    public static void main(String[] args) {
        String excelFilePath = "D:\\Documents\\Diploma\\TEMPLATE\\student_template.xlsx";
        String wordTemplatePath = "D:\\Documents\\Diploma\\TEMPLATE\\certificate_template.docx";
        String outputFolder = "D:\\Documents\\Diploma\\TEMPLATE\\certificates_try\\";

        try {
            File folder = new File(outputFolder);
            if (!folder.exists()) {
                folder.mkdirs();
            }

            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next(); // Skip header row

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                String studentName = getCellValueAsString(row.getCell(0));
                String registrationNo = getCellValueAsString(row.getCell(1));
                String collegeName = getCellValueAsString(row.getCell(2));
                String internshipDomain = getCellValueAsString(row.getCell(3));
                String issuanceDate = getCellValueAsString(row.getCell(4));
                String fromDate = getCellValueAsString(row.getCell(5));
                String toDate = getCellValueAsString(row.getCell(6));

                logger.info("Processing certificate for {}", studentName);

                Map<String, String> placeholders = new HashMap<>();
                placeholders.put("{{student_name}}", studentName);
                placeholders.put("{{registration_no}}", registrationNo);
                placeholders.put("{{college_name}}", collegeName);
                placeholders.put("{{Internship_area}}", internshipDomain);
                placeholders.put("{{issuance_date}}", issuanceDate);
                placeholders.put("{{from_date}}", fromDate);
                placeholders.put("{{to_date}}", toDate);

                generateCertificate(wordTemplatePath, outputFolder + studentName + "_certificate.docx", placeholders);
            }

            workbook.close();
            System.out.println("Certificates generated successfully!");

        } catch (IOException e) {
            logger.error("An error occurred: ", e);
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
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            replacePlaceholdersInParagraph(paragraph, placeholders);
        }

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
            StringBuilder fullText = new StringBuilder();
            List<XWPFRun> runs = paragraph.getRuns();

            for (XWPFRun run : runs) {
                fullText.append(run.text());
            }

            String updatedText = fullText.toString();
            boolean replaced = false;

            for (Map.Entry<String, String> entry : placeholders.entrySet()) {
                String placeholder = entry.getKey();
                if (updatedText.contains(placeholder)) {
                    updatedText = updatedText.replace(placeholder, entry.getValue());
                    replaced = true;
                }
            }

            if (replaced) {
                for (int i = runs.size() - 1; i >= 0; i--) {
                    paragraph.removeRun(i);
                }
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(updatedText);
         //       newRun.setBold(true);
            }
        } catch (Exception e) {
            logger.error("Error processing paragraph: ", e);
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return formatDate(cell.getDateCellValue());
                } else {
                    return (cell.getNumericCellValue() % 1 == 0) ?
                            String.valueOf((int) cell.getNumericCellValue()) :
                            String.valueOf(cell.getNumericCellValue()).trim();
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
        if (date == null) return "";
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd MMMM yyyy");
        return date.toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate().format(formatter);
    }
}
