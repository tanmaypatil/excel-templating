package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for JXLS template processing with phones using jx:each
 */
class PersonWithPhoneTemplateTest {

    private static final String TEMPLATE_PATH = "/person_template_with_phones.xlsx";

    @TempDir
    Path tempDir;

    @Test
    void testPersonWithPhonesAndAddress() throws IOException {
        // Given: A person with address and multiple phones
        Person person = new Person("John Doe", 15, "Jane Doe");

        Address address = new Address("Home", "123 Main St, Hometown");
        person.setAddress(address);

        person.addPhone("Mobile", "+1-555-1234");
        person.addPhone("Home", "+1-555-5678");
        person.addPhone("Work", "+1-555-9012");

        String outputPath = tempDir.resolve("person_with_phones_output.xlsx").toString();

        // When: Process the template
        processTemplate(TEMPLATE_PATH, outputPath, person);

        // Then: Output file should be generated
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists(), "Output Excel file should be created");
        assertTrue(outputFile.length() > 0, "Output file should not be empty");

        System.out.println("✓ Template with phones processed successfully: " + outputPath);
        System.out.println("  Person: " + person.getName() + ", Age: " + person.getAge());
        System.out.println("  Phones: " + person.getPhones().size());
        System.out.println("  Please open the Excel file to verify the output");
    }

    @Test
    void testAdultWithoutPhones() throws IOException {
        // Given: An adult without phones (no parent, no phones shown)
        Person person = new Person("Alice Smith", 25, "Bob Smith");

        String outputPath = tempDir.resolve("adult_without_phones_output.xlsx").toString();

        // When: Process the template
        processTemplate(TEMPLATE_PATH, outputPath, person);

        // Then: Output file should be generated
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists(), "Output Excel file should be created");

        System.out.println("✓ Adult without phones processed successfully: " + outputPath);
    }

    @Test
    void testMinorWithOnlyOnePhone() throws IOException {
        // Given: A minor with only one phone
        Person person = new Person("Tom Brown", 12, "Sarah Brown");
        person.addPhone("Mobile", "+1-555-1111");

        String outputPath = tempDir.resolve("minor_one_phone_output.xlsx").toString();

        // When: Process the template
        processTemplate(TEMPLATE_PATH, outputPath, person);

        // Then: Verify file is created
        assertTrue(new File(outputPath).exists(), "Output should be created");

        System.out.println("✓ Minor with one phone processed successfully");
    }

    @Test
    void testPersonWithManyPhones() throws IOException {
        // Given: Person with many phones
        Person person = new Person("Emma Wilson", 30, "David Wilson");

        Address address = new Address("Office", "456 Business Ave");
        person.setAddress(address);

        person.addPhone("Mobile", "+1-555-0001");
        person.addPhone("Home", "+1-555-0002");
        person.addPhone("Work", "+1-555-0003");
        person.addPhone("Fax", "+1-555-0004");
        person.addPhone("Emergency", "+1-555-0005");

        String outputPath = "/Users/tanmaypatil/excel-templating/person_with_many_phones_output.xlsx";

        // When: Process the template
        processTemplate(TEMPLATE_PATH, outputPath, person);

        // Then: Verify file is created
        File outputFile = new File(outputPath);
        assertTrue(outputFile.exists(), "Output should be created");
        assertTrue(outputFile.length() > 0, "Output should not be empty");

        System.out.println("✓ Person with many phones processed successfully: " + outputPath);
        System.out.println("  Person has " + person.getPhones().size() + " phone numbers");
        System.out.println("  Please open the Excel file to verify all phones are listed");
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
     * Verify the output contains correct data
     */
    private void verifyOutput(String outputPath, Person person) throws IOException {
        try (FileInputStream fis = new FileInputStream(outputPath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Find and verify name
            String nameValue = findCellValue(sheet, "Name:");
            assertNotNull(nameValue, "Name should be present");
            assertTrue(nameValue.contains(person.getName()),
                "Should contain name: " + person.getName());

            // Find and verify age
            String ageValue = findCellValue(sheet, "Age:");
            assertNotNull(ageValue, "Age should be present");
            assertTrue(ageValue.contains(String.valueOf(person.getAge())),
                "Should contain age: " + person.getAge());

            // If person has phones, verify they appear
            if (person.isPhoneExists()) {
                boolean foundPhoneNumbers = containsText(sheet, "Phone Numbers:");
                assertTrue(foundPhoneNumbers, "Should contain 'Phone Numbers:' section");

                // Verify each phone appears
                for (Phone phone : person.getPhones()) {
                    boolean foundPhoneType = containsText(sheet, phone.getPhoneType());
                    boolean foundPhoneNo = containsText(sheet, phone.getPhoneNo());

                    assertTrue(foundPhoneType,
                        "Should contain phone type: " + phone.getPhoneType());
                    assertTrue(foundPhoneNo,
                        "Should contain phone number: " + phone.getPhoneNo());
                }

                System.out.println("  ✓ All " + person.getPhones().size() + " phones verified in output");
            }
        }
    }

    /**
     * Find cell value in the next column after finding a label
     */
    private String findCellValue(Sheet sheet, String label) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING &&
                    cell.getStringCellValue().contains(label)) {
                    Cell nextCell = row.getCell(cell.getColumnIndex() + 1);
                    if (nextCell != null) {
                        return getCellValueAsString(nextCell);
                    }
                }
            }
        }
        return null;
    }

    /**
     * Check if sheet contains specific text
     */
    private boolean containsText(Sheet sheet, String text) {
        for (Row row : sheet) {
            if (row == null) continue;
            for (Cell cell : row) {
                if (cell == null) continue;
                String cellValue = getCellValueAsString(cell);
                if (cellValue.contains(text)) {
                    return true;
                }
            }
        }
        return false;
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
}
