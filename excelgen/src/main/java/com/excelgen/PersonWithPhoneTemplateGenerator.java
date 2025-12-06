package com.excelgen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Creates a proper JXLS 2.x template for Person with Address and Phones
 * Demonstrates use of jx:each for iterating over phone list
 */
public class PersonWithPhoneTemplateGenerator {

    public static void main(String[] args) throws IOException {
        String templatePath = "src/main/resources/person_template_with_phones.xlsx";
        createTemplateWithPhones(templatePath);
        System.out.println("Template with phones created at: " + templatePath);
    }

    public static void createTemplateWithPhones(String filePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("PersonInfo");

        // Styles
        CellStyle boldStyle = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldStyle.setFont(boldFont);

        // Drawing patriarch for comments
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper factory = workbook.getCreationHelper();

        // Row 0: Title with jx:area command in comment
        Row row0 = sheet.createRow(0);
        Cell titleCell = row0.createCell(0);
        titleCell.setCellValue("Person Report with Address and Phones");
        titleCell.setCellStyle(boldStyle);

        // Add jx:area command to cell A1 as a comment
        ClientAnchor areaAnchor = factory.createClientAnchor();
        areaAnchor.setCol1(0);
        areaAnchor.setCol2(3);
        areaAnchor.setRow1(0);
        areaAnchor.setRow2(2);
        Comment areaComment = drawing.createCellComment(areaAnchor);
        areaComment.setString(factory.createRichTextString("jx:area(lastCell=\"B15\")"));
        titleCell.setCellComment(areaComment);

        // Row 1: Empty
        sheet.createRow(1);

        // Row 2: Name
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Name:");
        row2.createCell(1).setCellValue("${person.name}");

        // Row 3: Age
        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("Age:");
        row3.createCell(1).setCellValue("${person.age}");

        // Row 4: Empty
        sheet.createRow(4);

        // Row 5: jx:if for parent (age < 18) - command in COMMENT
        Row row5 = sheet.createRow(5);
        Cell ifParentCell = row5.createCell(0);
        ifParentCell.setCellValue("Parent:");

        ClientAnchor anchor1 = factory.createClientAnchor();
        anchor1.setCol1(0);
        anchor1.setCol2(3);
        anchor1.setRow1(5);
        anchor1.setRow2(7);
        Comment comment1 = drawing.createCellComment(anchor1);
        comment1.setString(factory.createRichTextString("jx:if(condition=\"person.age < 18\", lastCell=\"B5\")"));
        ifParentCell.setCellComment(comment1);

        row5.createCell(1).setCellValue("${person.parentName}");

        // Row 6: Empty
        sheet.createRow(6);

        // Row 7: jx:if for address (addressExists is true) - command in COMMENT
        Row row7 = sheet.createRow(7);
        Cell ifAddressCell = row7.createCell(0);
        ifAddressCell.setCellValue("Address Type:");

        ClientAnchor anchor2 = factory.createClientAnchor();
        anchor2.setCol1(0);
        anchor2.setCol2(3);
        anchor2.setRow1(7);
        anchor2.setRow2(9);
        Comment comment2 = drawing.createCellComment(anchor2);
        comment2.setString(factory.createRichTextString("jx:if(condition=\"person.addressExists\", lastCell=\"B8\")"));
        ifAddressCell.setCellComment(comment2);

        row7.createCell(1).setCellValue("${person.address.type}");

        // Row 8: Address Line
        Row row8 = sheet.createRow(8);
        row8.createCell(0).setCellValue("Address:");
        row8.createCell(1).setCellValue("${person.address.addressLine}");

        // Row 9: Empty
        sheet.createRow(9);

        // Row 10: Phone Numbers header (static, no command)
        Row row10 = sheet.createRow(10);
        row10.createCell(0).setCellValue("Phone Numbers:");
        row10.getCell(0).setCellStyle(boldStyle);

        // Row 11 (POI index) = Row 12 in Excel: Phone data row with jx:each - command in COMMENT
        Row row11 = sheet.createRow(11);
        Cell phoneDataCell = row11.createCell(0);
        phoneDataCell.setCellValue("${phone.phoneType}");

        // Add jx:each command for iterating over phones (no jx:if wrapper)
        // IMPORTANT: Use Excel row number (12) not POI index (11)
        ClientAnchor anchor3 = factory.createClientAnchor();
        anchor3.setCol1(0);
        anchor3.setCol2(3);
        anchor3.setRow1(11);
        anchor3.setRow2(14);
        Comment comment3 = drawing.createCellComment(anchor3);
        comment3.setString(factory.createRichTextString("jx:each(items=\"person.phones\" var=\"phone\" lastCell=\"B12\" direction=\"DOWN\")"));
        phoneDataCell.setCellComment(comment3);

        row11.createCell(1).setCellValue("${phone.phoneNo}");

        // Row 13-15: Empty rows
        sheet.createRow(13);
        sheet.createRow(14);
        sheet.createRow(15);

        // Adjust column widths
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 8000);

        // Write to file
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }

        workbook.close();

        System.out.println("\nTemplate structure:");
        System.out.println("Row 1 (A1): Title + jx:area command in cell comment");
        System.out.println("Row 3: Name: ${person.name}");
        System.out.println("Row 4: Age: ${person.age}");
        System.out.println("Row 6: Parent: ${person.parentName} + jx:if command in cell comment");
        System.out.println("Row 8: Address Type: ${person.address.type} + jx:if command in cell comment");
        System.out.println("Row 9: Address: ${person.address.addressLine}");
        System.out.println("Row 11: Phone Numbers header + jx:if command in cell comment");
        System.out.println("Row 12: ${phone.phoneType} | ${phone.phoneNo} + jx:each command in cell comment (repeated for each phone)");
        System.out.println("\nIMPORTANT: Commands are in Excel cell COMMENTS, not cell values!");
        System.out.println("Open the template in Excel and hover over cells with red triangles to see commands.");
    }
}
