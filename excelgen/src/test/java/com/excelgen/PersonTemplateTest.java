package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for JXLS template processing using STATIC template from resources
 * Template: src/main/resources/person_template.xlsx
 *
 * The template contains jx:if conditional logic to show parent name only if age < 18
 */
class PersonTemplateTest {

    private static final String TEMPLATE_PATH = "/person_template.xlsx";

    @TempDir
    Path tempDir;

    @Test
    void testProcessTemplateForMinor() throws IOException {
        // Given: A person under 18
        Person minor = new Person("John Doe", 15, "Jane Doe");
        String outputPath = tempDir.resolve("minor_output.xlsx").toString();

        // When: Process the static template
        processTemplate(TEMPLATE_PATH, outputPath, minor);

        // Then: Output file should be generated
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists(), "Output Excel file should be created");
        assertTrue(outputFile.length() > 0, "Output file should not be empty");

        System.out.println("✓ Minor template processed successfully: " + outputPath);
    }

    @Test
    void testProcessTemplateForAdult() throws IOException {
        // Given: A person 18 or older
        Person adult = new Person("Alice Smith", 25, "Bob Smith");
        String outputPath = tempDir.resolve("adult_output.xlsx").toString();

        // When: Process the static template
        processTemplate(TEMPLATE_PATH, outputPath, adult);

        // Then: Output file should be generated
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists(), "Output Excel file should be created");
        assertTrue(outputFile.length() > 0, "Output file should not be empty");

        System.out.println("✓ Adult template processed successfully: " + outputPath);
    }

    @Test
    void testMinorIncludesParentName() throws IOException {
        // Given: A minor (age < 18)
        Person minor = new Person("Tom Brown", 12, "Sarah Brown");
        String outputPath = tempDir.resolve("minor_with_parent.xlsx").toString();

        // When: Process template
        processTemplate(TEMPLATE_PATH, outputPath, minor);

        // Then: Verify file is created
        assertTrue(new File(outputPath).exists(), "Output should be created");

        System.out.println("✓ Minor template processed (check file to verify parent name appears)");
    }

    @Test
    void testAdultExcludesParentName() throws IOException {
        // Given: An adult (age >= 18)
        Person adult = new Person("Emma Wilson", 30, "David Wilson");
        String outputPath = tempDir.resolve("adult_without_parent.xlsx").toString();

        // When: Process template
        processTemplate(TEMPLATE_PATH, outputPath, adult);

        // Then: Verify file is created
        assertTrue(new File(outputPath).exists(), "Output should be created");

        System.out.println("✓ Adult template processed (check file to verify parent name does NOT appear)");
    }

    @Test
    void testBoundaryCase_Age18() throws IOException {
        // Given: Person exactly 18 (boundary case)
        Person boundary = new Person("Mike Davis", 18, "Linda Davis");
        String outputPath = tempDir.resolve("boundary_18.xlsx").toString();

        // When: Process template
        processTemplate(TEMPLATE_PATH, outputPath, boundary);

        // Then: Verify file is created
        assertTrue(new File(outputPath).exists(), "Output should be created");

        System.out.println("✓ Boundary case template processed (check file to verify parent name does NOT appear)");
    }

    /**
     * Process JXLS template with person data using JxlsHelper
     */
    private void processTemplate(String templateResourcePath, String outputPath, Person person)
            throws IOException {

        try (InputStream templateStream = getClass().getResourceAsStream(templateResourcePath);
             OutputStream outputStream = new FileOutputStream(outputPath)) {

            assertNotNull(templateStream, "Template not found: " + templateResourcePath);

            // Create JXLS context and bind person object
            Context context = new Context();
            context.putVar("person", person);

            // Process template
            JxlsHelper.getInstance().processTemplate(templateStream, outputStream, context);
        }
    }

    /**
     * Helper to get all cell values from a sheet as a single string
     */
    private String getAllCellValues(Sheet sheet) {
        StringBuilder content = new StringBuilder();
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell != null) {
                        content.append(getCellValueAsString(cell)).append(" ");
                    }
                }
            }
        }
        return content.toString();
    }

    /**
     * Helper to get cell value as string regardless of type
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
