# JXLS Template Debugging Guide

## Problem
Template expressions like `${person.name}` not being replaced in production - headers intact but data missing.

## Quick Fix: Save to Temp File in Controller

```java
@PostMapping("/generate-report")
public ResponseEntity<byte[]> generateReport(@RequestBody PersonRequest request) {
    try {
        // Create person object
        Person person = new Person();
        person.setName(request.getName());
        person.setAge(request.getAge());

        // Load template
        InputStream templateStream = getClass().getClassLoader()
            .getResourceAsStream("person_template.xlsx");

        // Save to temp file for debugging
        File tempFile = new File(System.getProperty("java.io.tmpdir"),
            "debug_output_" + System.currentTimeMillis() + ".xlsx");
        System.out.println("DEBUG: Saving to: " + tempFile.getAbsolutePath());

        try (FileOutputStream fileOut = new FileOutputStream(tempFile)) {
            Context context = new Context();
            context.putVar("person", person);

            // Debug logging
            System.out.println("Context: " + context.toMap());
            System.out.println("Person: " + person);

            // Process template
            JxlsHelper.getInstance().processTemplate(templateStream, fileOut, context);
        }

        // Read and return
        byte[] fileContent = Files.readAllBytes(tempFile.toPath());

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDisposition(ContentDisposition.attachment()
            .filename("PersonReport.xlsx").build());

        return ResponseEntity.ok().headers(headers).body(fileContent);

    } catch (Exception e) {
        e.printStackTrace();
        return ResponseEntity.status(500).body(null);
    }
}
```

**Now open the temp file in Excel to see what's wrong!**

## Unit Test to Verify JXLS Processing

```java
@Test
public void testTemplateProcessing() throws IOException {
    // Create test person
    Person person = new Person();
    person.setName("John Doe");
    person.setAge(25);
    person.setParentName("Jane Doe");

    // Load template
    InputStream template = getClass().getClassLoader()
        .getResourceAsStream("person_template.xlsx");
    assertNotNull("Template not found!", template);

    // Process to file
    File output = new File("target/test-output.xlsx");
    try (FileOutputStream out = new FileOutputStream(output)) {
        Context context = new Context();
        context.putVar("person", person);
        JxlsHelper.getInstance().processTemplate(template, out, context);
    }

    System.out.println("Output saved to: " + output.getAbsolutePath());

    // Verify by reading back
    try (FileInputStream fis = new FileInputStream(output);
         Workbook wb = new XSSFWorkbook(fis)) {

        Sheet sheet = wb.getSheetAt(0);
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String value = cell.getStringCellValue();
                    System.out.println("Cell " + cell.getAddress() + ": " + value);

                    // Fail if expression not replaced
                    assertFalse("Expression not replaced: " + value,
                        value.contains("${"));
                }
            }
        }
    }
}
```

## Common Issues Checklist

1. **Template loading**
   ```java
   // Use this (no leading slash)
   getClass().getClassLoader().getResourceAsStream("person_template.xlsx");
   ```

2. **Context variable name must match template**
   ```java
   // Template has ${person.name}, so:
   context.putVar("person", person);  // ✓ Correct
   context.putVar("data", person);    // ✗ Wrong!
   ```

3. **Template has jx:area command in cell A1 comment**
   - Open template in Excel
   - Hover over cell A1 - should see red triangle
   - Comment should contain: `jx:area(lastCell="B10")`

4. **jx:if commands in cell COMMENTS, not values**
   - Cell value: `Parent:` (display text)
   - Cell comment: `jx:if(condition="person.age < 18", lastCell="B5")`

## Inspect Binary Excel File

**Method 1: Unzip the XLSX**
```bash
cp downloaded_file.xlsx downloaded_file.zip
unzip downloaded_file.zip -d excel_contents
cat excel_contents/xl/worksheets/sheet1.xml | grep '\${'
```
If you see `${person.name}` in the XML, JXLS didn't process it!

**Method 2: Read with POI**
```java
try (FileInputStream fis = new FileInputStream("output.xlsx");
     Workbook wb = new XSSFWorkbook(fis)) {
    Sheet sheet = wb.getSheetAt(0);
    for (Row row : sheet) {
        for (Cell cell : row) {
            if (cell.getCellType() == CellType.STRING) {
                String value = cell.getStringCellValue();
                if (value.contains("${")) {
                    System.err.println("Unprocessed: " + value);
                }
            }
        }
    }
}
```

## About the "author" Tag

Cell comments in Excel have an `author` attribute - this is normal and doesn't affect JXLS processing. The important part is the comment **content** (the jx:area or jx:if command).

## Dependencies Required

```xml
<dependency>
    <groupId>org.jxls</groupId>
    <artifactId>jxls</artifactId>
    <version>2.10.0</version>
</dependency>
<dependency>
    <groupId>org.jxls</groupId>
    <artifactId>jxls-poi</artifactId>
    <version>2.10.0</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>
```
