package com.rwto.excel.jexcelapi;

import com.rwto.excel.constant.ExcelConstant;
import jxl.*;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.File;

public class WriteExcelDemo {
    public static void main(String[] args) {
        try {
            String filePath = ExcelConstant.EXCEL_PATH_XLS;
            // 1. 创建工作簿
            WritableWorkbook workbook = Workbook.createWorkbook(new File(filePath));

            // 2. 创建工作表
            WritableSheet sheet = workbook.createSheet("Sheet1", 0);

            // 3. 定义单元格颜色
            WritableCellFormat greenFormat = new WritableCellFormat();
            greenFormat.setBackground(jxl.format.Colour.GREEN);

            WritableCellFormat yellowFormat = new WritableCellFormat();
            yellowFormat.setBackground(jxl.format.Colour.YELLOW);
            /**
             * -Xms64m -Xmx64m
             * row = 4000 col = 50
             * 正常写
             */
            // 4. 写入数据
            for (int row = 0; row < 4000; row++) {
                for (int col = 0; col < 50; col++) {
                    if(col%2 == 0){
                        Label label = new Label(col, row, "Cell " + (row + 1) + "-" + (col + 1),yellowFormat);
                        sheet.addCell(label);
                    }else{
                        Label label = new Label(col, row, "Cell " + (row + 1) + "-" + (col + 1),greenFormat);
                        sheet.addCell(label);
                    }
                }
            }

            // 5. 保存工作簿
            workbook.write();
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
