package com.test;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;

/**
 * Update CSL test case
 * @author I307451
 */
public class App 
{
    private static final Logger logger = LoggerFactory.getLogger(App.class);

    private static final String ROOT_PATH= System.getProperty("user.dir") + "\\in\\";

    public static void main( String[] args ) throws IOException {
        logger.warn("Current path: " + ROOT_PATH);
        logger.warn("********* Start Processing **********");

        //get to be updated cell list
        ArrayList<CSLCell> list = POIUtil.getRequiredCells();

        //test case list
        ArrayList<String> listFileName = new ArrayList<>();
        POIUtil.getAllFileName(ROOT_PATH,listFileName);
        for(String name:listFileName){
            if(name.contains(".xlsx")){
                logger.warn(name);
                POIUtil.updateCSL(name,list);
            }
        }

        logger.warn("********* Process End **********");
    }
}
