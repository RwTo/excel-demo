package com.rwto.excel.jexcelapi;

import com.rwto.excel.constant.ExcelConstant;
import jxl.*;

public class ReadExcelDemo {
    public static void main(String[] args) {
        try {
            // 1. 打开 Excel 文件
            String filePath = ExcelConstant.EXCEL_PATH_XLS;
            Workbook workbook = Workbook.getWorkbook(new java.io.File(filePath));

            // 2. 获取第一个工作表
            Sheet sheet = workbook.getSheet(0);
            /**
             * -Xms64m -Xmx64m
             * row = 4000 col = 50
             * 正常读
             */
            // 3. 遍历每一行，并读取数据
            for (int row = 0; row < sheet.getRows(); row++) {
                for (int col = 0; col < sheet.getColumns(); col++) {
                    Cell cell = sheet.getCell(col, row);
                    String content = cell.getContents();
                    content += cell.getCellFormat().getBackgroundColour().getDescription();
                    System.out.print(content+"\t");
                }
                System.out.println();
            }

            // 4. 关闭工作簿
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
