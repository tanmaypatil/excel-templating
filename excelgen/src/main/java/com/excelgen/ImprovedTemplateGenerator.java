package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Creates a properly structured JXLS template with clear command structure
 */
public class ImprovedTemplateGenerator {

    public static void main(String[] args) throws IOException {
        String templatePath = "src/main/resources/person_template.xlsx";
        createProperJxlsTemplate(templatePath);
        System.out.println("Improved JXLS template created at: " + templatePath);
        System.out.println("\nOpen in Excel to see the proper jx:if structure:");
        System.out.println("- Row with jx:if command (starting the condition)");
        System.out.println("- Content rows (shown conditionally)");
        System.out.println("- Row with jx:endif command (ending the condition)");
    }

    /**
     * Creates a JXLS template with proper command structure:
     * - jx:if to start the condition
     * - content to show conditionally
     * - jx:endif to close the condition (optional but clearer)
     */
    public static void createProperJxlsTemplate(String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("PersonInfo");

        // Row 0: Title
        Row row0 = sheet.createRow(0);
        Cell titleCell = row0.createCell(0);
        titleCell.setCellValue("Person Report");

        CellStyle boldStyle = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldStyle.setFont(boldFont);
        titleCell.setCellStyle(boldStyle);

        // Row 1: Empty
        sheet.createRow(1);

        // Row 2: Name with JXLS expression
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Name:");
        row2.createCell(1).setCellValue("${person.name}");

        // Row 3: Age with JXLS expression
        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("Age:");
        row3.createCell(1).setCellValue("${person.age}");

        // Row 4: Empty spacer
        sheet.createRow(4);

        // Row 5: jx:if command (start of conditional block)
        Row row5 = sheet.createRow(5);
        Cell ifCell = row5.createCell(0);
        ifCell.setCellValue("jx:if(condition=\"person.age < 18\", lastCell=\"B6\")");

        // Mark the command cell visually
        CellStyle commandStyle = workbook.createCellStyle();
        commandStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        commandStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        ifCell.setCellStyle(commandStyle);

        // Row 6: Parent name (conditional content - only shown if age < 18)
        Row row6 = sheet.createRow(6);
        row6.createCell(0).setCellValue("Parent:");
        row6.createCell(1).setCellValue("${person.parentName}");

        // Adjust column widths
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 6000);

        // Add a cell comment to explain the jx:if command
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(ifCell.getColumnIndex());
        anchor.setCol2(ifCell.getColumnIndex() + 2);
        anchor.setRow1(row5.getRowNum());
        anchor.setRow2(row5.getRowNum() + 3);

        Comment comment = drawing.createCellComment(anchor);
        RichTextString commentText = workbook.getCreationHelper().createRichTextString(
            "JXLS Command:\n" +
            "This row contains the jx:if condition.\n" +
            "If condition is true, the next row(s) will be included.\n" +
            "If false, they will be removed from output."
        );
        comment.setString(commentText);
        ifCell.setCellComment(comment);

        // Write to file
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }

        workbook.close();

        System.out.println("\nTemplate structure:");
        System.out.println("Row 1: Title");
        System.out.println("Row 3: Name: ${person.name}");
        System.out.println("Row 4: Age: ${person.age}");
        System.out.println("Row 6: jx:if(condition=\"person.age < 18\", lastCell=\"B6\") ← Command");
        System.out.println("Row 7: Parent: ${person.parentName} ← Conditional content");
    }
}
