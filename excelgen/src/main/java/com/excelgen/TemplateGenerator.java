package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Utility to generate the static JXLS template file.
 * Run this once to create the template, then use the static template.
 */
public class TemplateGenerator {

    public static void main(String[] args) throws IOException {
        String templatePath = "src/main/resources/person_template.xlsx";
        createPersonTemplate(templatePath);
        System.out.println("Static template created at: " + templatePath);
        System.out.println("\nThis template can now be used as a static resource.");
        System.out.println("You can open it in Excel and see the JXLS expressions.");
    }

    /**
     * Creates a static JXLS template with conditional parent name display
     *
     * Template Structure:
     * Row 1: "Person Report" (title)
     * Row 2: (empty)
     * Row 3: "Name: ${person.name}"
     * Row 4: "Age: ${person.age}"
     * Row 5: jx:if command (condition check)
     * Row 6: "Parent: ${person.parentName}" (conditional row)
     */
    public static void createPersonTemplate(String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("PersonInfo");

        // Row 0: Title
        Row row0 = sheet.createRow(0);
        Cell titleCell = row0.createCell(0);
        titleCell.setCellValue("Person Report");

        // Make title bold
        CellStyle boldStyle = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldStyle.setFont(boldFont);
        titleCell.setCellStyle(boldStyle);

        // Row 1: Empty for spacing
        sheet.createRow(1);

        // Row 2: Name with JXLS expression
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Name: ${person.name}");

        // Row 3: Age with JXLS expression
        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("Age: ${person.age}");

        // Row 4: jx:if command - this controls whether next row is shown
        Row row4 = sheet.createRow(4);
        Cell commandCell = row4.createCell(0);
        commandCell.setCellValue("jx:if(condition=\"person.age < 18\" lastCell=\"A5\" areas=[\"A5:A5\"])");

        // Style for command (light gray to indicate it's a command)
        CellStyle commandStyle = workbook.createCellStyle();
        commandStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        commandStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        commandCell.setCellStyle(commandStyle);

        // Row 5: Parent name - only shown if age < 18
        Row row5 = sheet.createRow(5);
        row5.createCell(0).setCellValue("Parent: ${person.parentName}");

        // Adjust column width for better readability
        sheet.setColumnWidth(0, 8000);

        // Write to file
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }

        workbook.close();
    }
}
