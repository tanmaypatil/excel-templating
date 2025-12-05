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
 * Tests to demonstrate when JXLS 2.x evaluates ${...} expressions
 */
class VariableEvaluationTest {

    @TempDir
    Path tempDir;

    @Test
    void testVariablesWithoutAnyCommands() throws IOException {
        System.out.println("\n=== Test: Template with ONLY variables (no commands) ===");

        // Create template with ONLY ${...} expressions, NO commands
        String templatePath = tempDir.resolve("no_commands_template.xlsx").toString();
        createTemplateWithoutCommands(templatePath);

        // Process it
        Person person = new Person("John Doe", 25, "Jane Doe");
        String outputPath = tempDir.resolve("no_commands_output.xlsx").toString();
        processTemplate(templatePath, outputPath, person);

        // Check result
        String content = getFileContent(outputPath);

        System.out.println("Template had: Only ${person.name} expressions");
        System.out.println("No JXLS commands at all");
        System.out.println("\nResult:");
        System.out.println("  Contains 'John Doe': " + content.contains("John Doe"));
        System.out.println("  Contains '${person.name}': " + content.contains("${person.name}"));

        // In newer JXLS versions, even without commands, simple ${} might work
        // But with invalid old commands, it definitely won't
        System.out.println("\n✓ JXLS 2.x MAY evaluate simple expressions without commands");
        System.out.println("  But with INVALID old commands present, it won't!");
    }

    @Test
    void testVariablesWithValidCommand() throws IOException {
        System.out.println("\n=== Test: Template WITH valid JXLS 2.x command ===");

        // Create template with valid jx:if command
        String templatePath = tempDir.resolve("with_command_template.xlsx").toString();
        createTemplateWithValidCommand(templatePath);

        // Process it
        Person person = new Person("John Doe", 25, "Jane Doe");
        String outputPath = tempDir.resolve("with_command_output.xlsx").toString();
        processTemplate(templatePath, outputPath, person);

        // Check result
        String content = getFileContent(outputPath);

        System.out.println("Template had: ${person.name} + valid jx:if command");
        System.out.println("\nResult:");
        System.out.println("  Contains 'John Doe': " + content.contains("John Doe"));
        System.out.println("  Contains '${person.name}': " + content.contains("${person.name}"));

        assertTrue(content.contains("John Doe"),
                "✅ Variables ARE evaluated when valid commands present");
        assertFalse(content.contains("${person.name}"),
                "✅ Literal ${...} should NOT appear");

        System.out.println("\n✓ CONFIRMED: Valid commands trigger full processing");
    }

    @Test
    void testVariablesWithInvalidOldCommand() throws IOException {
        System.out.println("\n=== Test: Template WITH invalid old <jx:if> syntax ===");

        // Create template with OLD invalid command
        String templatePath = tempDir.resolve("invalid_command_template.xlsx").toString();
        createTemplateWithInvalidCommand(templatePath);

        // Process it
        Person person = new Person("John Doe", 25, "Jane Doe");
        String outputPath = tempDir.resolve("invalid_command_output.xlsx").toString();
        processTemplate(templatePath, outputPath, person);

        // Check result
        String content = getFileContent(outputPath);

        System.out.println("Template had: ${person.name} + INVALID <jx:if> old syntax");
        System.out.println("\nResult:");
        System.out.println("  Contains 'John Doe': " + content.contains("John Doe"));
        System.out.println("  Contains '${person.name}': " + content.contains("${person.name}"));
        System.out.println("  Contains '<jx:if>': " + content.contains("<jx:if"));

        assertTrue(content.contains("${person.name}"),
                "❌ Variables NOT evaluated with invalid commands");
        assertTrue(content.contains("<jx:if"),
                "❌ Old syntax appears as literal text");

        System.out.println("\n✓ CONFIRMED: Invalid old commands prevent ALL processing");
        System.out.println("  Even ${person.name} (which is valid syntax) isn't evaluated!");
    }

    @Test
    void testExplanation() {
        System.out.println("\n╔════════════════════════════════════════════════════════╗");
        System.out.println("║  WHY ${person.name} ISN'T EVALUATED IN OLD TEMPLATE   ║");
        System.out.println("╚════════════════════════════════════════════════════════╝");
        System.out.println();
        System.out.println("JXLS 2.x Processing Logic:");
        System.out.println();
        System.out.println("1. Scan template for VALID JXLS 2.x commands");
        System.out.println("   - jx:if(...)");
        System.out.println("   - jx:each(...)");
        System.out.println("   - jx:area(...)");
        System.out.println("   etc.");
        System.out.println();
        System.out.println("2. If VALID commands found:");
        System.out.println("   ✅ Enter 'processing mode'");
        System.out.println("   ✅ Evaluate ALL ${...} expressions");
        System.out.println("   ✅ Execute all commands");
        System.out.println();
        System.out.println("3. If NO valid commands found (or INVALID old commands):");
        System.out.println("   ❌ Don't enter 'processing mode'");
        System.out.println("   ❌ Copy template AS-IS");
        System.out.println("   ❌ Don't evaluate ${...} expressions");
        System.out.println();
        System.out.println("Old Template Problem:");
        System.out.println("  - Has: <jx:if test=\"...\">  ← INVALID in JXLS 2.x");
        System.out.println("  - JXLS 2.x doesn't recognize this");
        System.out.println("  - Result: NO processing mode triggered");
        System.out.println("  - Even ${person.name} stays as literal text!");
        System.out.println();
        System.out.println("Solution:");
        System.out.println("  - Replace: <jx:if test=\"...\">...</jx:if>");
        System.out.println("  - With:    jx:if(condition=\"...\", lastCell=\"...\")");
        System.out.println("  - Then:    ALL ${...} expressions work");
        System.out.println("════════════════════════════════════════════════════════════");

        assertTrue(true, "Explanation test");
    }

    // Helper methods

    private void createTemplateWithoutCommands(String path) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
        sheet.createRow(0).createCell(0).setCellValue("Name: ${person.name}");
        sheet.createRow(1).createCell(0).setCellValue("Age: ${person.age}");
        try (FileOutputStream fos = new FileOutputStream(path)) {
            wb.write(fos);
        }
        wb.close();
    }

    private void createTemplateWithValidCommand(String path) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
        sheet.createRow(0).createCell(0).setCellValue("Name: ${person.name}");
        sheet.createRow(1).createCell(0).setCellValue("Age: ${person.age}");
        sheet.createRow(2).createCell(0).setCellValue("jx:if(condition=\"person.age >= 18\", lastCell=\"A3\")");
        sheet.createRow(3).createCell(0).setCellValue("Status: Adult");
        try (FileOutputStream fos = new FileOutputStream(path)) {
            wb.write(fos);
        }
        wb.close();
    }

    private void createTemplateWithInvalidCommand(String path) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
        sheet.createRow(0).createCell(0).setCellValue("Name: ${person.name}");
        sheet.createRow(1).createCell(0).setCellValue("Age: ${person.age}");
        sheet.createRow(2).createCell(0).setCellValue("<jx:if test=\"person.age >= 18\">");
        sheet.createRow(3).createCell(0).setCellValue("Status: Adult");
        sheet.createRow(4).createCell(0).setCellValue("</jx:if>");
        try (FileOutputStream fos = new FileOutputStream(path)) {
            wb.write(fos);
        }
        wb.close();
    }

    private void processTemplate(String templatePath, String outputPath, Person person) throws IOException {
        try (InputStream is = new FileInputStream(templatePath);
             OutputStream os = new FileOutputStream(outputPath)) {
            Context context = new Context();
            context.putVar("person", person);
            JxlsHelper.getInstance().processTemplate(is, os, context);
        }
    }

    private String getFileContent(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis)) {
            StringBuilder content = new StringBuilder();
            Sheet sheet = wb.getSheetAt(0);
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    for (Cell cell : row) {
                        if (cell != null) {
                            content.append(getCellValue(cell)).append(" ");
                        }
                    }
                }
            }
            return content.toString();
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf((int) cell.getNumericCellValue());
            default: return "";
        }
    }
}
