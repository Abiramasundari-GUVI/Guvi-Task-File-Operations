package com.file.operations.GuviFileOperationsTask;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {
    public static void main(String[] args) {
        // Create a workbook and a sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Data to write
        String[][] data = {
            {"Name", "Age", "Email"},
            {"John Doe", "30", "john@test.com"},
            {"Jane Doe", "28", "jane@test.com"},
            {"Bob Smith", "35", "jacky@example.com"},
            {"Swapnil", "37", "swapnil@example.com"}
        };

        // Populate the sheet with data
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < data[i].length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(data[i][j]);
            }
        }

        // Write to file
        try (FileOutputStream fos = new FileOutputStream("data.xlsx")) {
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Excel file written successfully.");
    }
}

