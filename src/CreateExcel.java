import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateExcel {

    static Row row;

    public static void main(String[] args) {
        try {
            FileInputStream docxFile = new FileInputStream("sample-document.docx");  // Make sure you entered poems Word file name correctly
            XWPFDocument document = new XWPFDocument(docxFile);

            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Poems Data");

            int rowNum = 0;
            boolean lineShiftAllowed = true;
            StringBuilder textBuilder = new StringBuilder();
            for (int i = 0; i < document.getParagraphs().size(); i++) {
                XWPFParagraph paragraph = document.getParagraphs().get(i);
                String text = paragraph.getText().trim();
                if (!text.isEmpty()) {
                    if (lineShiftAllowed) {
                        textBuilder = new StringBuilder();
                        row = sheet.createRow(rowNum++);
                        Cell cell1 = row.createCell(0);
                        Cell cell2 = row.createCell(1);
                        XWPFParagraph nextParagraph = document.getParagraphs().get(i + 1);
                        for (XWPFRun run : document.getParagraphs().get(i + 1).getRuns()) {
                            if (run.isBold()) {
                                cell2.setCellValue(nextParagraph.getText().trim());
                                i++;
                                break;
                            }
                        }

                        cell1.setCellValue(text);

                        lineShiftAllowed = false;
                    } else {
                        textBuilder.append(text).append("\n");
                    }
                } else if (i < document.getParagraphs().size() - 1) {
                    if (!document.getParagraphs().get(i + 1).isEmpty()) {
                        for (XWPFRun run : document.getParagraphs().get(i + 1).getRuns()) {
                            if (run.isBold()) {
                                lineShiftAllowed = true;
                                if (row == null) {
                                    row = sheet.createRow(0);
                                }
                                Cell cell3 = row.createCell(2);
                                cell3.setCellValue(textBuilder.toString().trim());
                                break;
                            }
                        }
                    }
                }
            }

            // Auto-size columns for better readability (optional)
            for (int i = 0; i < 3; i++) {
                sheet.autoSizeColumn(i);
            }

            // Save the Excel file
            FileOutputStream excelFile = new FileOutputStream("poems.xlsx");
            workbook.write(excelFile);
            excelFile.close();

            System.out.println("Excel file 'poems.xlsx' created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

