package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Verify that the generated Excel file has properly replaced expressions
 */
class VerifyOutputTest {

    @Test
    void verifyAddressOutputHasReplacedExpressions() throws IOException {
        String outputPath = "/Users/tanmaypatil/excel-templating/minor_with_address_output.xlsx";

        try (FileInputStream fis = new FileInputStream(outputPath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Verify name was replaced (should contain "John Doe", not "${person.name}")
            String nameValue = getCellValue(sheet, 2, 1);
            System.out.println("Name: " + nameValue);
            assertFalse(nameValue.contains("${"), "Name should not contain ${} expression");
            assertTrue(nameValue.contains("John Doe"), "Should contain actual name 'John Doe'");

            // Verify age was replaced
            String ageValue = getCellValue(sheet, 3, 1);
            System.out.println("Age: " + ageValue);
            assertFalse(ageValue.contains("${"), "Age should not contain ${} expression");
            assertTrue(ageValue.contains("15"), "Should contain actual age '15'");

            // Verify parent name was replaced (conditional - should appear for minor)
            String parentValue = getCellValue(sheet, 5, 1);
            System.out.println("Parent: " + parentValue);
            assertFalse(parentValue.contains("${"), "Parent should not contain ${} expression");
            assertTrue(parentValue.contains("Jane Doe"), "Should contain actual parent name 'Jane Doe'");

            // Verify address type was replaced
            String addressTypeValue = getCellValue(sheet, 7, 1);
            System.out.println("Address Type: " + addressTypeValue);
            assertFalse(addressTypeValue.contains("${"), "Address type should not contain ${} expression");
            assertTrue(addressTypeValue.contains("Home"), "Should contain actual address type 'Home'");

            // Verify address line was replaced
            String addressLineValue = getCellValue(sheet, 8, 1);
            System.out.println("Address Line: " + addressLineValue);
            assertFalse(addressLineValue.contains("${"), "Address line should not contain ${} expression");
            assertTrue(addressLineValue.contains("123 Main St"), "Should contain actual address");

            System.out.println("\n✓ All expressions were properly replaced!");
        }
    }

    private String getCellValue(Sheet sheet, int rowNum, int colNum) {
        Row row = sheet.getRow(rowNum);
        if (row == null) return "";

        Cell cell = row.getCell(colNum);
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
}
