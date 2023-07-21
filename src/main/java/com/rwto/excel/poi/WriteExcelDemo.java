package com.rwto.excel.poi;

import com.rwto.excel.constant.ExcelConstant;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class WriteExcelDemo {
    public static void main(String[] args) {
        try {

            String filePath = ExcelConstant.EXCEL_PATH_XLS;
            //String filePath = ExcelConstant.EXCEL_PATH_XLSX;
            //Workbook workbook = new XSSFWorkbook(); //支持xlsx文件的写入
            Workbook workbook = new HSSFWorkbook(); //支持xls文件的写入
            Sheet sheet = workbook.createSheet("Sheet1");
            /**
             * -Xms64m -Xmx64m
             * row = 4000 col = 50
             * 出现OOM
             */
            for (int row = 0; row < 4000; row++) {
                Row excelRow = sheet.createRow(row);
                for (int col = 0; col < 50; col++) {
                    Cell cell = excelRow.createCell(col);
                    cell.setCellValue("Cell " + (row + 1) + "-" + (col + 1));
                }
            }

            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);

            fileOutputStream.close();
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
