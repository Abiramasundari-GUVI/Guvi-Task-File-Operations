package com.file.operations.GuviFileOperationsTask;

import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.IOException;

public class ReadExcel {
    public static void main(String[] args) {
        try (FileInputStream fis = new FileInputStream("data.xlsx")) {
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheet("Sheet1");

            // Iterate through rows and cells
            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        default:
                            System.out.print("UNKNOWN\t");
                            break;
                    }
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
