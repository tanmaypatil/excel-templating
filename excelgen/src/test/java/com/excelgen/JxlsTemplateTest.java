package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * JUnit tests demonstrating JXLS template processing with jx:if conditional logic
 */
class JxlsTemplateTest {

    @TempDir
    Path tempDir;

    private String templatePath;

    @BeforeEach
    void createTemplate() throws IOException {
        templatePath = tempDir.resolve("person_template.xlsx").toString();

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("PersonInfo");

        // Row 0: Title
        Row row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("Person Report");

        // Row 1: Empty
        sheet.createRow(1);

        // Row 2: Name field
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Name: ${person.name}");

        // Row 3: Age field
        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("Age: ${person.age}");

        // Row 4: jx:if command - conditional rendering
        Row row4 = sheet.createRow(4);
        row4.createCell(0).setCellValue("jx:if(condition=\"person.age < 18\" lastCell=\"A5\" areas=[\"A5:A5\"])");

        // Row 5: Parent name (only shown if age < 18)
        Row row5 = sheet.createRow(5);
        row5.createCell(0).setCellValue("Parent: ${person.parentName}");

        try (FileOutputStream fos = new FileOutputStream(templatePath)) {
            wb.write(fos);
        }
        wb.close();

        System.out.println("Template created at: " + templatePath);
    }

    @Test
    void testMinorShowsParentName() throws IOException {
        // Given: A person under 18
        Person minor = new Person("John Doe", 15, "Jane Doe");
        String outputPath = tempDir.resolve("output_minor.xlsx").toString();

        // When: Process template
        processTemplate(templatePath, outputPath, minor);

        // Then: Verify file was created and contains data
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists(), "Output file should exist");
        assertTrue(outputFile.length() > 0, "Output file should not be empty");

        // Verify the content
        try (FileInputStream fis = new FileInputStream(outputPath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);

            // Verify name
            String nameValue = sheet.getRow(2).getCell(0).getStringCellValue();
            assertTrue(nameValue.contains("John Doe"), "Should contain person name");

            // Verify age
            String ageValue = sheet.getRow(3).getCell(0).getStringCellValue();
            assertTrue(ageValue.contains("15"), "Should contain person age");

            // Verify parent name is present (because age < 18)
            Row parentRow = sheet.getRow(4);
            if (parentRow != null && parentRow.getCell(0) != null) {
                String parentValue = parentRow.getCell(0).getStringCellValue();
                assertTrue(parentValue.contains("Jane Doe"),
                        "Should contain parent name for minor");
            }
        }

        System.out.println("Test passed: Minor shows parent name");
    }

    @Test
    void testAdultDoesNotShowParentName() throws IOException {
        // Given: A person over 18
        Person adult = new Person("Alice Smith", 25, "Bob Smith");
        String outputPath = tempDir.resolve("output_adult.xlsx").toString();

        // When: Process template
        processTemplate(templatePath, outputPath, adult);

        // Then: Verify file was created
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists(), "Output file should exist");
        assertTrue(outputFile.length() > 0, "Output file should not be empty");

        // Verify the content
        try (FileInputStream fis = new FileInputStream(outputPath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);

            // Verify name
            String nameValue = sheet.getRow(2).getCell(0).getStringCellValue();
            assertTrue(nameValue.contains("Alice Smith"), "Should contain person name");

            // Verify age
            String ageValue = sheet.getRow(3).getCell(0).getStringCellValue();
            assertTrue(ageValue.contains("25"), "Should contain person age");

            // Verify parent row either doesn't exist or is empty (because age >= 18)
            // The jx:if should have removed the parent row for adults
            int lastRowNum = sheet.getLastRowNum();
            assertTrue(lastRowNum <= 3,
                    "Should have fewer rows for adult (no parent row)");
        }

        System.out.println("Test passed: Adult does not show parent name");
    }

    @Test
    void testBoundaryCase_Exactly18() throws IOException {
        // Given: A person exactly 18 years old
        Person boundary = new Person("Boundary Person", 18, "Parent Name");
        String outputPath = tempDir.resolve("output_boundary.xlsx").toString();

        // When: Process template
        processTemplate(templatePath, outputPath, boundary);

        // Then: Verify parent name is NOT shown (condition is age < 18, not <=)
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists(), "Output file should exist");

        try (FileInputStream fis = new FileInputStream(outputPath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            assertTrue(lastRowNum <= 3,
                    "Should not show parent for person exactly 18 years old");
        }

        System.out.println("Test passed: Age 18 (boundary) does not show parent name");
    }

    /**
     * Helper method to process JXLS template
     */
    private void processTemplate(String templatePath, String outputPath, Person person) throws IOException {
        try (InputStream is = new FileInputStream(templatePath);
             OutputStream os = new FileOutputStream(outputPath)) {

            Context context = new Context();
            context.putVar("person", person);

            JxlsHelper.getInstance().processTemplate(is, os, context);
        }
    }
}
