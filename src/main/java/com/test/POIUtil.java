package com.test;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

public class POIUtil {

    private static final String C_SHEET_TEST_RESULT = "Test Cases and Results";
    private static final String C_SHEET_EVIDENCE = "Evidences";
    //C1 - 02
    private static final String C_H1 = "SAP Container Shipping Line Edition 2.1 Order to Cash";

    //F1 - 05
    private static final String C_H2 = "SAP Innovative Business Solutions\n" +
                                                    "Test Cases and Test Report";
    //CSL version - 14
    private static final String C_VERSION = "1.0";

    //Test Start Date - 24
    private static final String C_START_DATE = "November 20,2019";

    //Test End Date - 26
    private static final String C_END_DATE = "November 29,2019";

    //System Details - 42
    private static final String C_SYSTEM = "P4V 200";

    //Date - 62
    private static final String C_DATE1 = "November 15,2019";

    //Review Date - 64
    private static final String C_DATE2 = "November 17,2019";

    //Date - 66
    private static final String C_DATE3 = "November 19,2019";

    public static void updateCSL(String filePath) throws IOException {

        FileInputStream fsIP= new FileInputStream(new File(filePath)); //Read the spreadsheet that needs to be updated

        XSSFWorkbook wb = new XSSFWorkbook(fsIP); //Access the workbook

        //XSSFSheet worksheet = wb.getSheetAt(1); //Access the worksheet, so that we can update / modify it
        XSSFSheet worksheet = wb.getSheet(C_SHEET_TEST_RESULT);
        //create style
        XSSFCellStyle styleHeader = createStyle(wb);
        //Version
        updateCell(worksheet,1,4,C_VERSION,filePath,styleHeader);

        //C1 - 02 - header1 - "SAP Container Shipping Line Edition 2.1 Order to Cash"
        updateCell(worksheet,0,2,C_H1,filePath,null);

        //F1 - 05 -"SAP Innovative Business Solutions
        updateCell(worksheet,0,5,C_H2,filePath,null);

        //Test Start Date - 24 -"November 20,2019"
        updateCell(worksheet,2,4,C_START_DATE,filePath,styleHeader);

        //Test End Date - 26
        updateCell(worksheet,2,6,C_END_DATE,filePath,styleHeader);

        //System Details - 42 - "P4V 200";
        updateCell(worksheet,4,2,C_SYSTEM,filePath,styleHeader);

        //Date - 62 - "November 15,2019"
        updateCell(worksheet,6,2,C_DATE1,filePath,styleHeader);

        //Review Date - 64 - "November 17,2019"
        updateCell(worksheet,6,4,C_DATE2,filePath,styleHeader);

        //Date - 66 - "November 19,2019"
        updateCell(worksheet,6,6,C_DATE3,filePath,styleHeader);

//        XSSFCell cell = null; // declare a Cell object
//        //E2
//        cell = worksheet.getRow(1).getCell(4);   // Access the second cell in second row to update the value
//
//        if(cell == null){
//            System.out.println("File: " + filePath + " updated VERSION failed!");
//        }else{
//            //Version
//            cell.setCellValue(C_VERSION);  // Get current cell value value and overwrite the value
//        }

        //Clean up evidence
        cleanUpEvidence(wb);

        fsIP.close(); //Close the InputStream
        FileOutputStream output_file =new FileOutputStream(new File(filePath));  //Open FileOutputStream to write updates
        wb.write(output_file); //write changes
        output_file.close();  //close the stream
        System.out.println("File: " + filePath + " updated successfully!");
    }

    private static void cleanUpEvidence(XSSFWorkbook wb){
        XSSFSheet evidenceSheet = wb.getSheet(C_SHEET_EVIDENCE);
        if(evidenceSheet != null){
            for (int row = 0; row < 23; row++) {
                updateCell(evidenceSheet,row+3,3,"","filePath",null);
                updateCell(evidenceSheet,row+3,6,"","filePath",null);
            }
        }
    }

    /**
     * Update cell value
     * @param worksheet
     * @param row
     * @param col
     * @param value
     * @param path
     * @param style
     */
    private static void updateCell(XSSFSheet worksheet,int row, int col, String value,String path,CellStyle style){
        XSSFCell cell = null; // declare a Cell object
        cell = worksheet.getRow(row).getCell(col);
        if(cell == null){
            System.out.println("File: " + path + " updated VERSION failed!");
        }else{
            //Version
            cell.setCellValue(value);  // Get current cell value value and overwrite the value
            if(style != null){
                cell.setCellStyle(style);
            }
        }
    }
    /**
     * create style
     * @param wb
     */
    private static XSSFCellStyle createStyle(XSSFWorkbook wb){
        //创建一个字体
        Font font=wb.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Arial ");
//        font.setItalic(true);
//        font.setStrikeout(true);
        XSSFCellStyle style=wb.createCellStyle();

        style.setAlignment(HorizontalAlignment.LEFT); //字体右对齐
        style.setBorderBottom(BorderStyle.THIN);//下边框
//        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
//        style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
//        style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        style.setFont(font);
        return style;
    }
    /**
     * 获取一个文件夹下的所有文件全路径
     * @param path
     * @param listFileName
     */
    public static void getAllFileName(String path, ArrayList<String> listFileName){
        File file = new File(path);
        //listFiles()方法存储的是文件的完整路径
        File [] files = file.listFiles();
        String [] names = file.list();
        if(names != null){
            String [] completNames = new String[names.length];
            for(int i = 0;i < names.length;i++){
                completNames[i]=path+names[i];
            }
            listFileName.addAll(Arrays.asList(completNames));
        }
        for(File a:files){
            if(a.isDirectory()){//如果文件夹下有子文件夹，获取子文件夹下的所有文件全路径。
                if(!a.getName().contains("Archive")) {
                    getAllFileName(a.getAbsolutePath() + "\\", listFileName);
                }
            }
        }
    }
}
