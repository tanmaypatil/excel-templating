package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Creates a template using OLD JXLS 1.x XML-style syntax
 * This will NOT work with JXLS 2.x - for demonstration purposes only!
 */
public class OldStyleTemplateGenerator {

    public static void main(String[] args) throws IOException {
        String templatePath = "src/main/resources/person_template_old_style.xlsx";
        createOldStyleTemplate(templatePath);
        System.out.println("Old-style (JXLS 1.x) template created at: " + templatePath);
        System.out.println("\n⚠️  WARNING: This template uses JXLS 1.x syntax!");
        System.out.println("It will NOT work correctly with JXLS 2.x library.");
        System.out.println("This is for demonstration/testing purposes only.\n");
    }

    /**
     * Creates a template using JXLS 1.x XML-style syntax:
     * <jx:if test="condition">content</jx:if>
     *
     * This is the OLD syntax that does NOT work with JXLS 2.x!
     */
    public static void createOldStyleTemplate(String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("PersonInfo");

        // Row 0: Title
        Row row0 = sheet.createRow(0);
        Cell titleCell = row0.createCell(0);
        titleCell.setCellValue("Person Report (OLD JXLS 1.x Style)");

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

        // Row 5: OLD STYLE - Opening jx:if tag (JXLS 1.x)
        Row row5 = sheet.createRow(5);
        Cell oldStyleOpenTag = row5.createCell(0);
        oldStyleOpenTag.setCellValue("<jx:if test=\"person.age < 18\">");

        // Mark as old style with red background
        CellStyle oldStyleMarker = workbook.createCellStyle();
        oldStyleMarker.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        oldStyleMarker.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        oldStyleOpenTag.setCellStyle(oldStyleMarker);

        // Row 6: Conditional content
        Row row6 = sheet.createRow(6);
        row6.createCell(0).setCellValue("Parent:");
        row6.createCell(1).setCellValue("${person.parentName}");

        // Row 7: OLD STYLE - Closing jx:if tag (JXLS 1.x)
        Row row7 = sheet.createRow(7);
        Cell oldStyleCloseTag = row7.createCell(0);
        oldStyleCloseTag.setCellValue("</jx:if>");
        oldStyleCloseTag.setCellStyle(oldStyleMarker);

        // Row 8: Empty
        sheet.createRow(8);

        // Row 9: Explanation
        Row row9 = sheet.createRow(9);
        Cell explanation = row9.createCell(0);
        explanation.setCellValue("⚠️ This template uses OLD JXLS 1.x syntax (won't work with JXLS 2.x)");

        CellStyle warningStyle = workbook.createCellStyle();
        Font warningFont = workbook.createFont();
        warningFont.setColor(IndexedColors.RED.getIndex());
        warningStyle.setFont(warningFont);
        explanation.setCellStyle(warningStyle);

        // Adjust column widths
        sheet.setColumnWidth(0, 8000);
        sheet.setColumnWidth(1, 6000);

        // Add cell comment to explain
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(oldStyleOpenTag.getColumnIndex());
        anchor.setCol2(oldStyleOpenTag.getColumnIndex() + 2);
        anchor.setRow1(row5.getRowNum());
        anchor.setRow2(row5.getRowNum() + 3);

        Comment comment = drawing.createCellComment(anchor);
        RichTextString commentText = workbook.getCreationHelper().createRichTextString(
            "JXLS 1.x Syntax (OLD):\n" +
            "This XML-style syntax does NOT work with JXLS 2.x!\n\n" +
            "JXLS 2.x will treat these as literal text, not commands."
        );
        comment.setString(commentText);
        oldStyleOpenTag.setCellComment(comment);

        // Write to file
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }

        workbook.close();

        System.out.println("\nTemplate structure (OLD JXLS 1.x style):");
        System.out.println("Row 3: Name: ${person.name}");
        System.out.println("Row 4: Age: ${person.age}");
        System.out.println("Row 6: <jx:if test=\"person.age < 18\"> ← OLD syntax");
        System.out.println("Row 7: Parent: ${person.parentName}");
        System.out.println("Row 8: </jx:if> ← OLD syntax");
    }
}
