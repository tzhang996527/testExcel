package com.test;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;
import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public final class ExcelUtil {
    /**
     * 从Excel中读取文件，读取的文件是一个DTO类，该类必须继承BaseRowModel
     * 具体实例参考 ： MemberMarketDto.java
     * 参考：https://github.com/alibaba/easyexcel
     * 字符流必须支持标记，FileInputStream 不支持标记，可以使用BufferedInputStream 代替
     * BufferedInputStream bis = new BufferedInputStream(new FileInputStream(...));
     */
    public static <T extends BaseRowModel> List<T> readExcel(final InputStream inputStream, final Class<? extends BaseRowModel> clazz) {
        if (null == inputStream) {
            throw new NullPointerException("the inputStream is null!");
        }
        ExcelModelListener<T> listener = new ExcelModelListener<>();
        // 这里因为EasyExcel-1.1.1版本的bug，所以需要选用下面这个标记已经过期的版本
        ExcelReader reader = new ExcelReader(inputStream,  valueOf(inputStream), null, listener);
        reader.read(new com.alibaba.excel.metadata.Sheet(1, 1, clazz));

        return listener.getRows();
    }


    public static void writeExcel(final File file, List<? extends BaseRowModel> list) {
        try (OutputStream out = new FileOutputStream(file)) {
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX);
            //写第一个sheet,  有模型映射关系
            Class<? extends BaseRowModel> t = list.get(0).getClass();
            Sheet sheet = new Sheet(1, 0, t);
            writer.write(list, sheet);
            writer.finish();
        } catch (IOException e) {
            System.out.println("fail to write to excel file: file[{}]" + file.getName());
        }
    }


    /**
     * 根据输入流，判断为xls还是xlsx，该方法原本存在于easyexcel 1.1.0 的ExcelTypeEnum中。
     */
    public static ExcelTypeEnum valueOf(InputStream inputStream) {
        try {
            FileMagic fileMagic = FileMagic.valueOf(inputStream);
            if (FileMagic.OLE2.equals(fileMagic)) {
                return ExcelTypeEnum.XLS;
            }
            if (FileMagic.OOXML.equals(fileMagic)) {
                return ExcelTypeEnum.XLSX;
            }
            throw new IllegalArgumentException("excelTypeEnum can not null");

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void simpleWrite(){

        // 文件输出位置
        String outPath = "C:\\Users\\oukele\\Desktop\\test.xlsx";

        try {
            // 所有行的集合
            List<List<Object>> list = new ArrayList<List<Object>>();

            for (int i = 1; i <= 10; i++) {
                // 第 n 行的数据
                List<Object> row = new ArrayList<Object>();
                row.add("第" + i + "单元格");
                row.add("第" + i + "单元格");
                list.add(row);
            }

            ExcelWriter excelWriter = EasyExcelFactory.getWriter(new FileOutputStream(outPath));
            // 表单
            Sheet sheet = new Sheet(1,0);
            sheet.setSheetName("第一个Sheet");
            // 创建一个表格
            Table table = new Table(1);
            // 动态添加 表头 headList --> 所有表头行集合
            List<List<String>> headList = new ArrayList<List<String>>();
            // 第 n 行 的表头
            List<String> headTitle0 = new ArrayList<String>();
            List<String> headTitle1 = new ArrayList<String>();
            List<String> headTitle2 = new ArrayList<String>();
            headTitle0.add("最顶部-1");
            headTitle0.add("标题1");
            headTitle1.add("最顶部-1");
            headTitle1.add("标题2");
            headTitle2.add("最顶部-1");
            headTitle2.add("标题3");

            headList.add(headTitle0);
            headList.add(headTitle1);
            headList.add(headTitle2);
            table.setHead(headList);

            excelWriter.write1(list,sheet,table);
            // 记得 释放资源
            excelWriter.finish();
            System.out.println("ok");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }

    // 简单读取 (同步读取)
    public static void simpleRead() {

        // 读取 excel 表格的路径
        String readPath = "C:\\Users\\oukele\\Desktop\\模拟数据.xlsx";


        try {
            // sheetNo --> 读取哪一个 表单
            // headLineMun --> 从哪一行开始读取( 不包括定义的这一行，比如 headLineMun为2 ，那么取出来的数据是从 第三行的数据开始读取 )
            // clazz --> 将读取的数据，转化成对应的实体，需要 extends BaseRowModel
            Sheet sheet = new Sheet(1, 1, ExcelModel.class);

            // 这里 取出来的是 ExcelModel实体 的集合
            List<Object> readList = EasyExcelFactory.read(new FileInputStream(readPath), sheet);
            // 存 ExcelMode 实体的 集合
            List<ExcelModel> list = new ArrayList<ExcelModel>();
            for (Object obj : readList) {
                list.add((ExcelModel) obj);
            }

            // 取出数据
            StringBuilder str = new StringBuilder();
            str.append("{");
            String link = "";
            for (ExcelModel mode : list) {
                str.append(link).append("\""+mode.getName()+"\":").append("\""+mode.getNickName()+"\"");
                link= ",";
            }
            str.append("};");
            System.out.println(str);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    // 异步读取
    public static void simpleRead1(){
        // 读取 excel 表格的路径
        String readPath = "C:\\Users\\oukele\\Desktop\\模拟数据.xlsx";

        try {
            Sheet sheet = new Sheet(1,1,ExcelModel.class);
            EasyExcelFactory.readBySax(new FileInputStream(readPath),sheet,new ExcelModelListener());

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }


}
