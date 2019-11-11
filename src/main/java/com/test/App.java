package com.test;

import java.io.*;
import java.util.ArrayList;

/**
 * Hello world!
 *
 */
public class App 
{
    private static final String ROOT_PATH= System.getProperty("user.dir") + "\\in\\";

    public static void main( String[] args ) throws IOException {
        System.out.println("用户的当前工作目录:"+System.getProperty("user.dir") );

        //String rootPath = System.getProperty("user.dir") + "\\in\\";
        //String workPath = ROOT_PATH + "file_in.xlsx";;

//        try (FileInputStream inputStream = new FileInputStream(work_path)) {
//            List<ExcelModel> users = ExcelUtil.readExcel(new BufferedInputStream(inputStream), ExcelModel.class);
//            //System.out.println(users);
//            for(ExcelModel user: users){
//                System.out.println(user.getName() + "," + user.getNickName());
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }

        ArrayList<String> listFileName = new ArrayList<>();
        POIUtil.getAllFileName(ROOT_PATH,listFileName);
        for(String name:listFileName){
            if(name.contains(".xlsx")||name.contains(".properties")){
                System.out.println(name);
                POIUtil.updateCSL(name);
            }
        }

        System.out.println("********* Process End **********");
    }
}
