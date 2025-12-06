package com.excelgen;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;

public class ExcelDumper {
    public static void main(String[] args) throws Exception {
        String filePath = "/Users/tanmaypatil/excel-templating/person_with_many_phones_output.xlsx";

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            System.out.println("=== Excel File Content ===");
            System.out.println("File: " + filePath);
            System.out.println();

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    System.out.println("Row " + (i + 1) + ": (empty)");
                    continue;
                }

                StringBuilder sb = new StringBuilder();
                sb.append("Row ").append(i + 1).append(": ");

                for (int j = 0; j < 5; j++) {  // Check first 5 columns
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        String value = getCellValue(cell);
                        if (!value.isEmpty()) {
                            sb.append("|").append(value);
                        }
                    }
                }

                System.out.println(sb.toString());
            }
        }
    }

    private static String getCellValue(Cell cell) {
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
