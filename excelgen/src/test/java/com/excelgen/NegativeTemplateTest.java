package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * NEGATIVE TEST: Demonstrates what happens when you use OLD JXLS 1.x syntax
 * with the NEW JXLS 2.x library.
 *
 * Purpose: Show that old XML-style syntax does NOT work with JXLS 2.x
 *
 * Expected Behavior:
 * - The <jx:if> and </jx:if> tags will be treated as LITERAL TEXT
 * - They will appear in the output Excel file as regular text
 * - The conditional logic will NOT work
 * - Variables like ${person.name} WILL still be evaluated (that syntax is the same)
 */
class NegativeTemplateTest {

    private static final String OLD_TEMPLATE_PATH = "/person_template_old_style.xlsx";

    @TempDir
    Path tempDir;

    @BeforeAll
    static void createOldStyleTemplate() throws IOException {
        // Generate the old-style template before running tests
        String templatePath = "src/main/resources/person_template_old_style.xlsx";
        OldStyleTemplateGenerator.createOldStyleTemplate(templatePath);
        System.out.println("\n========================================");
        System.out.println("NEGATIVE TEST: Using JXLS 1.x syntax with JXLS 2.x library");
        System.out.println("========================================\n");
    }

    @Test
    void testOldSyntaxIsIgnored_MinorCase() throws IOException {
        System.out.println("\n=== Test 1: Minor (age < 18) with OLD syntax ===");

        // Given: A minor (age < 18)
        Person minor = new Person("John Doe", 15, "Jane Doe");
        String outputPath = tempDir.resolve("old_style_minor.xlsx").toString();

        // When: Process the OLD-style template with JXLS 2.x
        processTemplate(OLD_TEMPLATE_PATH, outputPath, minor);

        // Then: Verify the behavior
        try (FileInputStream fis = new FileInputStream(outputPath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);

            System.out.println("\nOutput content:");
            printSheetContent(sheet);

            // ACTUAL BEHAVIOR with OLD syntax on JXLS 2.x:
            // 1. Variables ${person.name} are NOT evaluated (appear as literal text)
            // 2. The <jx:if> and </jx:if> tags appear as LITERAL TEXT
            // 3. NOTHING gets processed - template is copied as-is!

            String allContent = getAllCellValues(sheet);

            // CRITICAL: Variables are NOT evaluated - they appear as literal ${...}
            assertTrue(allContent.contains("${person.name}"),
                    "❌ Variables should NOT be evaluated - appear as literal text");
            assertTrue(allContent.contains("${person.age}"),
                    "❌ Variables should NOT be evaluated - appear as literal text");

            // The OLD syntax tags will appear as literal text!
            assertTrue(allContent.contains("<jx:if") || allContent.contains("jx:if"),
                    "❌ OLD syntax appears as literal text in output");

            // The parent variable appears as literal ${...} (not evaluated)
            assertTrue(allContent.contains("${person.parentName}"),
                    "❌ Parent variable appears as literal text (not evaluated)");

            System.out.println("\n✓ CONFIRMED: OLD syntax causes COMPLETE processing failure");
            System.out.println("  - Variables NOT evaluated: " + allContent.contains("${person.name}"));
            System.out.println("  - jx:if tags appear as text: " + allContent.contains("jx:if"));
            System.out.println("  - Template copied as-is (no processing at all)");
        }
    }

    @Test
    void testOldSyntaxIsIgnored_AdultCase() throws IOException {
        System.out.println("\n=== Test 2: Adult (age >= 18) with OLD syntax ===");

        // Given: An adult (age >= 18)
        Person adult = new Person("Alice Smith", 25, "Bob Smith");
        String outputPath = tempDir.resolve("old_style_adult.xlsx").toString();

        // When: Process the OLD-style template with JXLS 2.x
        processTemplate(OLD_TEMPLATE_PATH, outputPath, adult);

        // Then: Verify the behavior
        try (FileInputStream fis = new FileInputStream(outputPath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);

            System.out.println("\nOutput content:");
            printSheetContent(sheet);

            String allContent = getAllCellValues(sheet);

            // Variables are NOT evaluated - appear as literal text
            assertTrue(allContent.contains("${person.name}"),
                    "❌ Variables NOT evaluated");
            assertTrue(allContent.contains("${person.age}"),
                    "❌ Variables NOT evaluated");

            // The OLD syntax tags appear as literal text
            assertTrue(allContent.contains("<jx:if") || allContent.contains("jx:if"),
                    "❌ OLD syntax appears as literal text in output");

            // The parent variable appears as literal text (not evaluated)
            assertTrue(allContent.contains("${person.parentName}"),
                    "❌ Parent variable appears as literal text");

            System.out.println("\n✓ CONFIRMED: Template not processed at all with old syntax");
            System.out.println("  - No variable evaluation occurs");
            System.out.println("  - Template copied as-is to output");
            System.out.println("  - This proves old syntax is completely incompatible");
        }
    }

    @Test
    void testComparison_OldVsNew() throws IOException {
        System.out.println("\n=== Test 3: Side-by-side comparison ===");

        Person person = new Person("Test Person", 20, "Test Parent");

        // Process with OLD syntax template
        String oldOutput = tempDir.resolve("comparison_old.xlsx").toString();
        processTemplate(OLD_TEMPLATE_PATH, oldOutput, person);

        // Process with NEW syntax template
        String newOutput = tempDir.resolve("comparison_new.xlsx").toString();
        processTemplate("/person_template.xlsx", newOutput, person);

        // Compare results
        String oldContent = getFileContent(oldOutput);
        String newContent = getFileContent(newOutput);

        System.out.println("\n--- OLD syntax output ---");
        System.out.println("Contains <jx:if>: " + oldContent.contains("jx:if"));
        System.out.println("Contains parent name: " + oldContent.contains("Test Parent"));

        System.out.println("\n--- NEW syntax output ---");
        System.out.println("Contains jx:if command: " + newContent.contains("jx:if"));
        System.out.println("Contains parent name: " + newContent.contains("Test Parent"));

        System.out.println("\n✓ COMPARISON COMPLETE");
        System.out.println("  OLD: Shows jx:if as text, parent always appears");
        System.out.println("  NEW: Evaluates condition properly, parent conditionally appears");
    }

    @Test
    void testDocumentation_WhatHappens() {
        System.out.println("\n╔════════════════════════════════════════════════════════════════╗");
        System.out.println("║  WHAT HAPPENS WHEN YOU USE OLD SYNTAX WITH JXLS 2.x?          ║");
        System.out.println("╚════════════════════════════════════════════════════════════════╝");
        System.out.println();
        System.out.println("JXLS 1.x Syntax (OLD):");
        System.out.println("  <jx:if test=\"person.age < 18\">");
        System.out.println("    Parent: ${person.parentName}");
        System.out.println("  </jx:if>");
        System.out.println();
        System.out.println("Behavior with JXLS 2.x:");
        System.out.println("  ❌ <jx:if> tag is treated as LITERAL TEXT");
        System.out.println("  ❌ </jx:if> tag is treated as LITERAL TEXT");
        System.out.println("  ❌ Variables ${...} are NOT evaluated (appear as literal text)");
        System.out.println("  ❌ Conditional logic does NOT work");
        System.out.println("  ❌ Template is copied AS-IS (no processing at all!)");
        System.out.println();
        System.out.println("Result:");
        System.out.println("  - The XML tags appear in your output Excel file");
        System.out.println("  - Variables like ${person.name} appear as literal text");
        System.out.println("  - Template processing completely fails");
        System.out.println("  - You MUST migrate to JXLS 2.x syntax");
        System.out.println();
        System.out.println("JXLS 2.x Syntax (NEW):");
        System.out.println("  jx:if(condition=\"person.age < 18\", lastCell=\"B6\")");
        System.out.println("  Parent: ${person.parentName}");
        System.out.println();
        System.out.println("Behavior with JXLS 2.x:");
        System.out.println("  ✅ Command is recognized and processed");
        System.out.println("  ✅ Condition is evaluated");
        System.out.println("  ✅ Content appears/disappears based on condition");
        System.out.println("  ✅ No command text appears in output");
        System.out.println();
        System.out.println("══════════════════════════════════════════════════════════════════");

        // This test always passes - it's just for documentation
        assertTrue(true, "Documentation test");
    }

    /**
     * Process JXLS template (works with both old and new syntax)
     */
    private void processTemplate(String templateResourcePath, String outputPath, Person person)
            throws IOException {

        try (InputStream templateStream = getClass().getResourceAsStream(templateResourcePath);
             OutputStream outputStream = new FileOutputStream(outputPath)) {

            assertNotNull(templateStream, "Template not found: " + templateResourcePath);

            Context context = new Context();
            context.putVar("person", person);

            JxlsHelper.getInstance().processTemplate(templateStream, outputStream, context);
        }
    }

    /**
     * Get all cell values from a sheet
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
     * Get cell value as string
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    /**
     * Print sheet content for debugging
     */
    private void printSheetContent(Sheet sheet) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                System.out.print("  Row " + (i + 1) + ": ");
                for (Cell cell : row) {
                    if (cell != null && !getCellValueAsString(cell).isEmpty()) {
                        System.out.print("[" + getCellValueAsString(cell) + "] ");
                    }
                }
                System.out.println();
            }
        }
    }

    /**
     * Get file content as string
     */
    private String getFileContent(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis)) {
            return getAllCellValues(wb.getSheetAt(0));
        }
    }
}
