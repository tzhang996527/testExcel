package com.test;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

public class POIUtil {

    private static final Logger logger = LoggerFactory.getLogger(App.class);

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
        if(worksheet == null){
            logger.warn("Error: " + filePath + " test sheet not found" );
            return;
        }
        //create style
        XSSFCellStyle styleHeader = createStyle(wb);
        //Version
        updateCell(worksheet,1,4,C_VERSION,"VERSION",styleHeader);

        //C1 - 02 - header1 - "SAP Container Shipping Line Edition 2.1 Order to Cash"
        updateCell(worksheet,0,2,C_H1,"H1",null);

        //F1 - 05 -"SAP Innovative Business Solutions
        updateCell(worksheet,0,5,C_H2,"H2",null);

        //Test Start Date - 24 -"November 20,2019"
        updateCell(worksheet,2,4,C_START_DATE,"START_DATE",styleHeader);

        //Test End Date - 26
        updateCell(worksheet,2,6,C_END_DATE,"C_END_DATE",styleHeader);

        //System Details - 42 - "P4V 200";
        updateCell(worksheet,4,2,C_SYSTEM,"C_SYSTEM",styleHeader);

        //Date - 62 - "November 15,2019"
        updateCell(worksheet,6,2,C_DATE1,"C_DATE1",styleHeader);

        //Review Date - 64 - "November 17,2019"
        updateCell(worksheet,6,4,C_DATE2,"C_DATE2",styleHeader);

        //Date - 66 - "November 19,2019"
        updateCell(worksheet,6,6,C_DATE3,"C_DATE3",styleHeader);

        //Clean up evidence
        cleanUpEvidence(wb);

        fsIP.close(); //Close the InputStream
        FileOutputStream output_file =new FileOutputStream(new File(filePath));  //Open FileOutputStream to write updates
        wb.write(output_file); //write changes
        output_file.close();  //close the stream
        logger.warn("File: " + filePath + " updated successfully!");
    }

    private static void cleanUpEvidence(XSSFWorkbook wb){
        String lv_msg;
        XSSFSheet evidenceSheet = wb.getSheet(C_SHEET_EVIDENCE);
        if(evidenceSheet != null){
            for (int row = 0; row < 10; row++) {
                lv_msg = "evidence " + Integer.toString(row);
                updateCell(evidenceSheet,row+3,3,"",lv_msg,null);
                updateCell(evidenceSheet,row+3,6,"",lv_msg,null);
            }
        }
    }

    /**
     * Update cell value
     * @param worksheet worksheet
     * @param row row
     * @param col col
     * @param value value
     * @param path file path
     * @param style style
     */
    private static void updateCell(XSSFSheet worksheet,int row, int col, String value,String path,CellStyle style){
        XSSFCell cell; // declare a Cell object
        cell = worksheet.getRow(row).getCell(col);
        if(cell == null){
            logger.warn("Error>>>> updated "+ path + " failed!");
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
     * @param wb workbook
     */
    private static XSSFCellStyle createStyle(XSSFWorkbook wb){
        //create font
        Font font=wb.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Arial ");
        XSSFCellStyle style=wb.createCellStyle();

        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderBottom(BorderStyle.THIN);
        style.setFont(font);
        return style;
    }
    /**
     * Get all files
     * @param path file path
     * @param listFileName all file names
     */
    public static void getAllFileName(String path, ArrayList<String> listFileName){
        File file = new File(path);
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
            if(a.isDirectory()){//Get all files
                if(!a.getName().contains("Archive")) { //skip archive folder
                    getAllFileName(a.getAbsolutePath() + "\\", listFileName);
                }
            }
        }
    }
}
