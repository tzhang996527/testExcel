package com.test;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;

import static org.apache.poi.ss.usermodel.Cell.*;

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

    //Evidence skip line
    private static final int C_SKIP_LINE = 3;
    private static final int C_SKIP_LINE2 = 10;

    /**
     * collect required cells
     *
     * @return List CSLCell list
     */
    public static ArrayList<CSLCell> getRequiredCells() {
        ArrayList<CSLCell> list = new ArrayList<>();

        //Version
        list.add(new CSLCell(1, 4, C_VERSION, "VERSION", "H"));

        //C1 - 02 - header1 - "SAP Container Shipping Line Edition 2.1 Order to Cash"
        list.add(new CSLCell(0, 2, C_H1, "H1", "N"));

        //F1 - 05 -"SAP Innovative Business Solutions
        list.add(new CSLCell(0, 5, C_H2, "C_H2", "N"));

        //Test Start Date - 24 -"November 20,2019"
        list.add(new CSLCell(2, 4, C_START_DATE, "START_DATE", "H"));

        //Test End Date - 26
        list.add(new CSLCell(2, 6, C_END_DATE, "C_END_DATE", "H"));

        //System Details - 42 - "P4V 200";
        list.add(new CSLCell(4, 2, C_SYSTEM, "C_SYSTEM", "H"));

        //Date - 62 - "November 15,2019"
        list.add(new CSLCell(6, 2, C_DATE1, "C_DATE1", "H"));

        //Review Date - 64 - "November 17,2019"
        list.add(new CSLCell(6, 4, C_DATE2, "C_DATE2", "H"));

        //Date - 66 - "November 19,2019"
        list.add(new CSLCell(6, 6, C_DATE3, "C_DATE3", "H"));

        return list;
    }

    /**
     * Update cell
     *
     * @param filePath input file
     * @throws IOException
     */
    public static void updateCSL(String filePath, ArrayList<CSLCell> cellArrayList) throws IOException {

        FileInputStream fsIP = new FileInputStream(new File(filePath)); //Read the spreadsheet that needs to be updated

        XSSFWorkbook wb = new XSSFWorkbook(fsIP); //Access the workbook

        //XSSFSheet worksheet = wb.getSheetAt(1); //Access the worksheet, so that we can update / modify it
        XSSFSheet worksheet = wb.getSheet(C_SHEET_TEST_RESULT);
        if (worksheet == null) {
            logger.warn("Error: " + filePath + " test sheet not found");
            return;
        }

        //create style
        XSSFCellStyle styleHeader = createStyle(wb);

        //Update cell
        for (CSLCell c : cellArrayList) {
            if (c.getStyle() == "H") {
                updateCell(worksheet, c, styleHeader);
            } else if (c.getStyle() == "N") {
                updateCell(worksheet, c, null);
            }
        }

        //clear evidence in test result sheet
        clearTestResult(worksheet,C_SKIP_LINE2,4);

        //Clean up evidence in evidence sheet
        XSSFSheet evidenceSheet = wb.getSheet(C_SHEET_EVIDENCE);
        if (evidenceSheet != null) {
            clearTestResult(evidenceSheet,C_SKIP_LINE,2);
        }

        fsIP.close(); //Close the InputStream
        FileOutputStream output_file = new FileOutputStream(new File(filePath));  //Open FileOutputStream to write updates
        wb.write(output_file); //write changes
        output_file.close();  //close the stream
        logger.warn("File: " + filePath + " updated successfully!");
    }

    /**
     * clear range content
     * @param sheet
     * @param skipRow
     * @param startCol
     */
    private static void clearTestResult(XSSFSheet sheet, int skipRow,int startCol) {
        DataFormatter df = new DataFormatter();
        boolean lv_need_clear;
        //test result starts from F11
        for (Iterator rowIterator = sheet.iterator(); rowIterator.hasNext(); ) {
            XSSFRow row = (XSSFRow) rowIterator.next();
            if (row.getRowNum() < skipRow) { //skip first 10 lines
                continue;
            }

            logger.debug("Row:>>>>>" + row.getRowNum());
            lv_need_clear = false;
            for (Iterator iterator = row.cellIterator(); iterator.hasNext(); ) {
                XSSFCell cell = (XSSFCell) iterator.next();
                String str = df.formatCellValue(cell);
                //Verify first column
                if(cell.getColumnIndex() == 0){
                    if(StringUtils.isNumeric(str)){
                        lv_need_clear = true;
                        continue;
                    }else{
                        break;
                    }
                }

                if(lv_need_clear && cell.getColumnIndex() > startCol){
                    cell.setCellValue((String) null);
                    cell.removeHyperlink();
//                    cell.setCellType(CellType.BLANK);
                    logger.debug("Cell cleared:->" + cell.getColumnIndex());
                }
            }

        }
    }

    /**
     * Update cell value
     *
     * @param worksheet worksheet
     * @param cslCell   cell object
     * @param style     style
     */
    private static void updateCell(XSSFSheet worksheet, CSLCell cslCell, CellStyle style) {
        XSSFCell cell; // declare a Cell object
        cell = worksheet.getRow(cslCell.getRow()).getCell(cslCell.getCol());
        if (cell == null) {
            logger.warn("Error>>>> updated "+ cslCell.getName() + " failed!");
        } else {
            //Version
            cell.setCellValue(cslCell.getValue());  // Get current cell value value and overwrite the value
            if (style != null) {
                cell.setCellStyle(style);
            }
        }
    }

    /**
     * create style
     *
     * @param wb workbook
     */
    private static XSSFCellStyle createStyle(XSSFWorkbook wb) {
        //create font
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Arial ");
        XSSFCellStyle style = wb.createCellStyle();

        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderBottom(BorderStyle.THIN);
        style.setFont(font);
        return style;
    }

    /**
     * Get all files
     *
     * @param path         file path
     * @param listFileName all file names
     */
    public static void getAllFileName(String path, ArrayList<String> listFileName) {
        File file = new File(path);
        File[] files = file.listFiles();
        String[] names = file.list();
        if (names != null) {
            String[] completNames = new String[names.length];
            for (int i = 0; i < names.length; i++) {
                completNames[i] = path + names[i];
            }
            listFileName.addAll(Arrays.asList(completNames));
        }
        for (File a : files) {
            if (a.isDirectory()) {//Get all files
                if (!a.getName().contains("Archive")) { //skip archive folder
                    getAllFileName(a.getAbsolutePath() + "\\", listFileName);
                }
            }
        }
    }
}
