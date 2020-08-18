package com.night.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;

/**
 * @Author: CharmNight
 * @Date: 2020/8/18 23:26
 */
public class ReturnClient {

    //响应到客户端
    public static void returnClient(HttpServletResponse response, String fileName, HSSFWorkbook wb) {
        try {
            setResponseHeader(response, fileName);
            OutputStream os = response.getOutputStream();
            wb.write(os);
            os.flush();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //发送响应流方法
    public static void setResponseHeader(HttpServletResponse response, String fileName) {
        try {
            try {
                fileName = new String(fileName.getBytes(), "ISO8859-1");
            } catch (UnsupportedEncodingException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            response.setContentType("application/octet-stream;charset=ISO8859-1");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
            response.addHeader("Pargam", "no-cache");
            response.addHeader("Cache-Control", "no-cache");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    //通过输出流响应
    public static void returnClient(String fileName, HSSFWorkbook wb) throws Exception {
        FileOutputStream os = null;
        try {
            File file = new File(fileName);
            os = new FileOutputStream(file);
            wb.write(os);
            os.flush();
            os.close();
        } catch (Exception e) {
            throw e;
        } finally {
            if (os != null) {
                try {
                    // 关闭流
                    os.close();
                } catch (IOException e) {
                    throw e;
                }
            }
        }
    }

}
