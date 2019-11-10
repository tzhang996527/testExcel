package com.test;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

public class POIUtil {
    public static void updateCSL(String filePath) throws IOException {

        FileInputStream fsIP= new FileInputStream(new File(filePath)); //Read the spreadsheet that needs to be updated

        XSSFWorkbook wb = new XSSFWorkbook(fsIP); //Access the workbook

        XSSFSheet worksheet = wb.getSheetAt(1); //Access the worksheet, so that we can update / modify it.

        XSSFCell cell = null; // declare a Cell object

        cell = worksheet.getRow(2).getCell(5);   // Access the second cell in second row to update the value

        cell.setCellValue("CSL2.2");  // Get current cell value value and overwrite the value

        fsIP.close(); //Close the InputStream

        FileOutputStream output_file =new FileOutputStream(new File(filePath));  //Open FileOutputStream to write updates

        wb.write(output_file); //write changes

        output_file.close();  //close the stream

        System.out.println("File: " + filePath + " updated successfully!");
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
                getAllFileName(a.getAbsolutePath()+"\\",listFileName);
            }
        }
    }
}
