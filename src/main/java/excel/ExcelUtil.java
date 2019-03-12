package excel;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @Discription
 * @Date 2019-03-05 17:09
 * @Author James
 */
public class ExcelUtil<T> {

    public static final String XLS = "xls";
    public static final String XLSX = "xlsx";
    public static final String CSV = "csv";


    /**
     * 解析获取 excel 并获取指定位置的值
     * @param is InputStream
     * @param isExcel2003 是否是 xls
     * @param location 需要取的数据的位置，第一个从 0 开始
     * @return
     */
    public static List<String> readExcelValue(InputStream is, boolean isExcel2003, Object[] location) {
        List<String> list = null;
        try {
            /** 根据版本选择创建Workbook的方式 */
            Workbook wb = null;
            // 当excel是2003时
            if (isExcel2003) {
                wb = new HSSFWorkbook(is);
            } else {// 当excel是2007时
                wb = new XSSFWorkbook(is);
            }
            // 读取Excel里面客户的信息
            list = readExcelValue(wb, location);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return list;

    }


    public static List<String> readExcelValue(Workbook wb, Object[] location) {
        // 总行数
        int totalRows = 0;
        // 总条数
        int totalCells = 0;

        // 得到第一个shell
        Sheet sheet = wb.getSheetAt(0);

        // 得到Excel的行数
        totalRows = sheet.getPhysicalNumberOfRows();

        // 得到Excel的列数(前提是有行数)
        if (totalRows >= 1 && sheet.getRow(0) != null) {
            totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }

        List<String> customerList = new ArrayList<String>();

        // 循环Excel行数,从第二行开始。标题不入库

        for (int r = 1; r < totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;

            String no = null;
            for (int c = 0; c < totalCells; c++) {
                row.getCell(c).setCellType(Cell.CELL_TYPE_STRING);
                Cell cell = row.getCell(c);
                for (int i = 0; i < location.length; i ++) {
                    if ((Integer) location[i] == c) {
                        customerList.add(cell.getStringCellValue());
                    }
                }
            }
        }
        return customerList;
    }

    /**
     * 获取行数
     * @param is
     * @param isExcel2003
     * @return
     */
    public static int getTotalRows(InputStream is, boolean isExcel2003) {
        /** 根据版本选择创建Workbook的方式 */
        Workbook wb = null;
        // 当excel是2003时
        try {
            if (isExcel2003) {
                wb = new HSSFWorkbook(is);
            } else {// 当excel是2007时
                wb = new XSSFWorkbook(is);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 得到第一个shell
        Sheet sheet = wb.getSheetAt(0);
        // 得到Excel的行数
        return sheet.getPhysicalNumberOfRows();
    }

}
