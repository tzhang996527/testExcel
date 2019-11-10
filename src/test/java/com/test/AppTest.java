package com.test;

import static org.junit.Assert.assertTrue;

import org.junit.Test;

/**
 * Unit test for simple App.
 */
public class AppTest 
{
    /**
     * Rigorous Test :-)
     */
    @Test
    public void shouldAnswerWithTrue()
    {
        assertTrue( true );
    }

    @Test
    public void simpleRead() {
        String fileName = System.getProperty("user.dir") + "\\in\\" + "file_in.xlsx";
        System.out.println(fileName);
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        //EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
    }

}
