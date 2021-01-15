package com.office.modules.sys.utils;/**
 * Created by wei on 2020/12/24 14:15
 */

import com.office.common.utils.FileComments;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLDecoder;
import java.net.URLEncoder;
import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletResponse;
import lombok.val;

/**
 * 下载
 * @Author wei
 * @Date 2020/12/24 14:15
 *
 **/
public class DownloadUtil {


    /**
     * 下载文件
     * @param response
     * @param filePath
     * @throws Exception
     */
    public static void downloadFile(HttpServletResponse response,String filePath) throws Exception{
        filePath=URLDecoder.decode(filePath,"utf-8");
        System.out.println(filePath);
        File file = new File(filePath);
        String fileName=file.getName();
        // 取得文件的后缀名。
        String ext = fileName.substring(fileName.lastIndexOf(".") + 1).toUpperCase();
        InputStream inStream = new FileInputStream(filePath);
        response.setContentType("multipart/form-data");
        response.setHeader("Content-Disposition", "attachment; filename=" +URLEncoder.encode(fileName,"utf-8"));
        response.setBufferSize(10240);
        response.setContentLength(Integer.valueOf(String.valueOf(file.length())));
        response.setCharacterEncoding("utf-8");
        // 循环取出流中的数据
        byte[] b = new byte[10240];
        int len;
        OutputStream out = response.getOutputStream();
        try {
            while ((len = inStream.read(b)) > 0){
                out.write(b, 0, len);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            inStream.close();
            out.close();
            out.flush();
        }
        System.out.println("success");
    }

    /**
     * 在线打开
     * @param filePath
     * @param response
     * @throws Exception
     */
    public static void inLineFile(HttpServletResponse response,String filePath)throws Exception{
        File f = new File(filePath);
        if (!f.exists()) {
            response.sendError(404, "File not found!");
            return;
        }
        BufferedInputStream br = new BufferedInputStream(new FileInputStream(f));
        byte[] buf = new byte[1024];
        int len = 0;
        response.reset(); // 非常重要
        URL u = new URL("file:///" + filePath);
        response.setContentType(u.openConnection().getContentType());
        response.setHeader("Content-Disposition", "inline; filename=" + f.getName());
        // 文件名应该编码成UTF-8
        OutputStream out = response.getOutputStream();
        while ((len = br.read(buf)) > 0)
            out.write(buf, 0, len);
        br.close();
        out.close();
    }
}
