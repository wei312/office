package com.office.modules.export.utils;/** 
 * Created by wei on 2020/12/17 16:30
 */

import static javax.swing.text.StyleConstants.ALIGN_CENTER;
import static org.apache.poi.hssf.record.cf.BorderFormatting.BORDER_THIN;
import static org.apache.poi.hssf.usermodel.HSSFCellStyle.*;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.helpers.ColumnShifter;
import org.springframework.lang.Nullable;

/**
 *
 * @Author wei
 * @Date 2020/12/17 16:30
 *
 **/
public class ExcelUtil{

    /*
     * 导出EXCEL列
     * @Author wei
     * @Date 16:47 2020/12/17
     * @Param [lists, columns, columnsName, outFilePath]
     * @return boolean
     **/
    public static HSSFWorkbook WriteToDocExcel(List<LinkedHashMap> lists,String[] columns,String[] columnsName )throws Exception{
        //创建工作薄对象
        HSSFWorkbook workbook=new HSSFWorkbook();//这里也可以设置sheet的Name
        //创建工作表对象
        HSSFSheet sheet = workbook.createSheet("sheet1");
        //创建工作表的行
        int i=0;
        HSSFRow row = sheet.createRow(i);
        for (int j = 0; j < columnsName.length; j++) {
            sheet.autoSizeColumn(j);
            HSSFCell cell =row.createCell(j);
            cell.setCellValue(columnsName[j].toString());
            addCellStyle(i,workbook,sheet,workbook.createCellStyle(),cell);
        }
        i++;
        for (int j = 0; j <lists.size(); j++) {
            LinkedHashMap linkedHashMap=lists.get(j);
            row = sheet.createRow(i);
            for (int k = 0; k < columns.length; k++) {
                sheet.autoSizeColumn(k);
                HSSFCell cell =row.createCell(k);
                cell.setCellValue(linkedHashMap.get(columns[k]).toString());
                addCellStyle(i,workbook,sheet,workbook.createCellStyle(),cell);
            }
            i++;
        }
        return workbook;
    }

    /**
     *
     * @param fileName
     * @param lists
     * @param columns
     * @param columnsName
     * @param filePath
     * @return
     * @throws Exception
     */
    public static String WriteToDocExcel(String fileName,List<LinkedHashMap> lists,String[] columns,String[] columnsName,String filePath )throws Exception{
        HSSFWorkbook workbook = WriteToDocExcel(lists,columns,columnsName);
        return exportData(workbook,fileName,filePath);
    }


        /**
         * 将导出excel写入response
         */
    public static String exportData(Workbook workbook,String fileName,String filePath) throws Exception {
        File file=new File(filePath);
        if(!file.exists()){
            file.createNewFile();
        }
        FileOutputStream fileOutputStream =new FileOutputStream(file);
        workbook.write(fileOutputStream);
        return filePath;
    }


    /**
     * 添加样式
     */
    public static void addCellStyle(int i,HSSFWorkbook workbook,HSSFSheet sheet,HSSFCellStyle cellStyle,HSSFCell cell){
        HSSFFont font = workbook.createFont();
        if(i==0){
            cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);// 设置背景色
            cellStyle.setFillPattern(SOLID_FOREGROUND);
            //设置字体:
            font.setFontName("黑体");
            font.setBold(true);//粗体显示
            font.setFontHeightInPoints((short) 14);//设置字体大小
        }
        //设置边框:
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        //设置居中:
        cellStyle.setAlignment(HorizontalAlignment.CENTER); // 居中
        // 设置字体
        font.setFontName("仿宋_GB2312");
        font.setFontHeightInPoints((short) 10);
        cellStyle.setFont(font);//选择需要用到的字体格式
        //设置自动换行:
        cellStyle.setWrapText(true);//设置自动换行
        cell.setCellStyle(cellStyle);
    }
}
