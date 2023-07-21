package com.rwto.excel.poi;

import com.rwto.excel.constant.ExcelConstant;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class SXSSFWriteExcelDemo {
    public static void main(String[] args) {
        try {
            // Set custom temporary directory
            String customTempDirPath = "temp";
            System.setProperty("java.io.tmpdir", customTempDirPath);

            String filePath = ExcelConstant.EXCEL_PATH_XLSX_BIG;
            SXSSFWorkbook workbook = new SXSSFWorkbook(500);//rowAccessWindowSize为内存中缓存的记录数,默认100
            Sheet sheet = workbook.createSheet("Sheet1");
            for (int row = 0; row < 100000; row++) {
                Row excelRow = sheet.createRow(row);
                for (int col = 0; col < 3; col++) {
                    Cell cell = excelRow.createCell(col);
                    cell.setCellValue("Cell " + (row + 1) + "-" + (col + 1));
                }
            }

            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);

            fileOutputStream.close();
            //使用 dispose() 方法释放（删除） SXSSFWorkbook 使用的临时资源，特别是在写入大量数据后，这一步骤很重要。
            workbook.dispose();
            /*如果不手动释放，默认等虚拟机停止也会删除临时文件
            poi.keep.tmp.files 通过这个配置可以控制虚拟机停止时，不删除临时文件*/

            //程序执行结束后，手动删除临时文件目录(看是否需要)
            deleteTempFiles(new File(customTempDirPath));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    private static void deleteTempFiles(File directory)  {
        if (directory.isDirectory()) {
            File[] files = directory.listFiles();
            if (files != null) {
                for (File file : files) {
                    deleteTempFiles(file);
                }
            }
        }
        if (!directory.delete()) {
            System.err.println("Failed to delete temp file: " + directory.getAbsolutePath());
        } else {
            System.out.println("Deleted temp file: " + directory.getAbsolutePath());
        }
    }

}
