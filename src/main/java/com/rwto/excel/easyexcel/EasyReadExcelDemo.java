package com.rwto.excel.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.rwto.excel.constant.ExcelConstant;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

public class EasyReadExcelDemo {

    public static void main(String[] args) {
        String filePath = ExcelConstant.EE_EXCEL_PATH_XLSX_BIG;

        // 使用 EasyExcel 读取 Excel 文件
        EasyExcel.read(filePath, UserData.class, new UserDataListener()).sheet().doRead();
    }

    public static class UserDataListener extends AnalysisEventListener<UserData> {
        private List<UserData> dataList = new ArrayList<>();

        @Override
        public void invoke(UserData data, AnalysisContext context) {
            dataList.add(data);
        }

        @Override
        public void doAfterAllAnalysed(AnalysisContext context) {
            // 在这里可以对 dataList 中的数据进行处理，比如保存到数据库或其他操作
            for (UserData userData : dataList) {
                System.out.println(userData.getName() + "\t" + userData.getAge() + "\t" + userData.getEmail());
            }
        }
    }
}
