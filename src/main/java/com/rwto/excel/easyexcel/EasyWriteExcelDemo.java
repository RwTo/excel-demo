package com.rwto.excel.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.rwto.excel.constant.ExcelConstant;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

public class EasyWriteExcelDemo {

    public static void main(String[] args) {
        String filePath = ExcelConstant.EE_EXCEL_PATH_XLSX_BIG;

        // 创建写入的数据列表
        List<UserData> dataList = new ArrayList<>();
        dataList.add(new UserData("Alice", 25, "alice@example.com"));
        dataList.add(new UserData("Bob", 30, "bob@example.com"));
        dataList.add(new UserData("Charlie", 28, "charlie@example.com"));

        // 使用 EasyExcel 写入 Excel 文件
        EasyExcel.write(filePath, UserData.class).sheet("Sheet1").doWrite(dataList);
    }
}
