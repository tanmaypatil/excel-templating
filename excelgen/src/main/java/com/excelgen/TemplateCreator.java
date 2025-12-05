package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class TemplateCreator {

    /**
     * Creates a simple JXLS template with conditional logic
     * Using jx:if command to conditionally display parent name
     */
    public static void createTemplate(String templatePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Persons");

        // Create header row
        Row headerRow = sheet.createRow(0);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("Person Information");

        // Empty row
        sheet.createRow(1);

        // Name field
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Name:");
        row2.createCell(1).setCellValue("${person.name}");

        // Age field
        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("Age:");
        row3.createCell(1).setCellValue("${person.age}");

        // Empty row for jx:if command
        Row row4 = sheet.createRow(4);
        Cell commandCell = row4.createCell(0);
        // JXLS command to conditionally show parent name if age < 18
        commandCell.setCellValue("jx:if(condition=\"person.age < 18\", lastCell=\"B5\", areas=[\"A6:B6\"])");

        // Conditional row - only displayed if age < 18
        Row row5 = sheet.createRow(5);
        row5.createCell(0).setCellValue("Parent Name:");
        row5.createCell(1).setCellValue("${person.parentName}");

        // Write to file
        try (FileOutputStream fileOut = new FileOutputStream(templatePath)) {
            workbook.write(fileOut);
        }
        workbook.close();

        System.out.println("Template created at: " + templatePath);
    }

    /**
     * Process the template with person data
     */
    public static void processTemplate(String templatePath, String outputPath, Person person) throws IOException {
        try (InputStream is = new FileInputStream(templatePath);
             OutputStream os = new FileOutputStream(outputPath)) {

            Context context = new Context();
            context.putVar("person", person);

            JxlsHelper.getInstance().processTemplate(is, os, context);

            System.out.println("Excel generated at: " + outputPath);
        }
    }
}
