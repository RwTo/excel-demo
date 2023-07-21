package com.rwto.excel.poi;

import com.rwto.excel.constant.ExcelConstant;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;

public class ReadExcelDemo {
    public static void main(String[] args) {
        try {
            String filePath = ExcelConstant.EXCEL_PATH_XLSX;
            //String filePath = ExcelConstant.EXCEL_PATH_XLSX;
            FileInputStream fileInputStream = new FileInputStream(filePath);
            //WorkbookFactory会根据文件类型自动选择使用HSSFWorkBook 或 XSSFWorkBook
            Workbook workbook = WorkbookFactory.create(fileInputStream);

            Sheet sheet = workbook.getSheetAt(0);
            /**
             * -Xms64m -Xmx64m
             * row = 4000 col = 50
             * 出现OOM
             */
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String cellValue = getCellValueAsString(cell);
                    System.out.print(cellValue + "\t");
                }
                System.out.println();
            }

            fileInputStream.close();
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }
}
