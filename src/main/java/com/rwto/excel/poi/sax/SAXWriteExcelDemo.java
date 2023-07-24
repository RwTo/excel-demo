package com.rwto.excel.poi.sax;

import com.rwto.excel.constant.ExcelConstant;
import lombok.Data;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

/**
 * @author rwto
 * @create 2023/7/20 14:29
 **/
public class SAXWriteExcelDemo {

    public static void main(String[] args) throws Exception {
        String filePath = ExcelConstant.EXCEL_PATH_XLSX_BIG;
        InputStream is = new FileInputStream(filePath);
        OPCPackage opcPackage = OPCPackage.open(is);
        try {
            // 读取 Excel 文件
            XSSFReader reader = new XSSFReader(opcPackage);
            // 使用事件处理器处理 Sheet 数据
            SAXSheetHandler sheetHandler = new SAXSheetHandler();
            XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(reader.getStylesTable(), reader.getSharedStringsTable(), sheetHandler, false);
            XMLReader sheetParser = XMLReaderFactory.createXMLReader();
            sheetParser.setContentHandler(xmlHandler);
            Iterator<InputStream> sheetsData = reader.getSheetsData();
            while(sheetsData.hasNext()){
                InputStream sheetIs = sheetsData.next();
                InputSource sheet = new InputSource(sheetIs);
                //开始解析sheet页
                System.out.println("=====================================开始处理sheet页==============================================");
                long start = System.currentTimeMillis();
                sheetParser.parse(sheet);
                System.out.println("=====================================处理sheet结束==============================================");
                long end = System.currentTimeMillis();
                List<List<SAXCell>> curSheet = sheetHandler.getCurSheet();
                curSheet.forEach(System.out::println);

                System.out.println("处理行数："+sheetHandler.getRowCount());
                System.out.println("处理单元格数："+sheetHandler.getCellCount());
                System.out.println("处理时间（ms）："+(end-start));

                sheetHandler.clear();
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (OpenXML4JException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        }finally {
            // 关闭输入流和 OPCPackage
            is.close();
            opcPackage.close();
        }
    }
}

class SAXSheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    //当前页
    private List<List<SAXCell>> curSheet;
    //当前行
    private List<SAXCell> curRow;
    //总单元格数
    private int cellCount;

    public SAXSheetHandler() {
        this.curSheet = new LinkedList<>();
    }

    /**
     * 一行处理开始时
     *
     * @param i 行号
     */
    @Override
    public void startRow(int i) {
        //开始处理新的一行，初始化当前行
        curRow = new LinkedList<>();
    }

    /**
     * 一行处理结束时
     *
     * @param i 行号
     */
    @Override
    public void endRow(int i) {
        //一行处理结束，将这一行数据存入sheet
        curSheet.add(curRow);
    }


    /**
     * 处理单元格(不会读取空单元格)
     *
     * @param cellReference  单元格名称
     * @param formattedValue 单元格值
     * @param xssfComment    批注
     */
    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment xssfComment) {
        SAXCell cell = new SAXCell(cellReference, formattedValue);
        curRow.add(cell);
        cellCount++;
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        // 头部和页脚
    }

    @Override
    public void endSheet() {
        //sheet处理结束
    }

    public List<List<SAXCell>> getCurSheet() {
        return curSheet;
    }

    public int getCellCount() {
        return cellCount;
    }

    public int getRowCount() {
        return curSheet.size();
    }

    public void clear(){
        curSheet = new LinkedList<>();
        cellCount = 0;
    }
}
@Data
class SAXCell {
    private String cellName;
    private String value;

    public SAXCell(String cellName, String value) {
        this.cellName = cellName;
        this.value = value;
    }
}
