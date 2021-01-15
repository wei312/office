package com.office.common.utils;/**
 * Created by wei on 2020/12/17 17:31
 */

import java.util.Locale;

/**
 *
 * @Author wei
 * @Date 2020/12/17 17:31
 *
 **/
public class FileComments {

    public static void main(String[] args) {
        System.out.println(System.getProperty("os.name"));
    }
    /*
     * 获取项目根路径
     * @Author wei
     * @Date 17:58 2020/12/17
     * @Param []
     * @return java.lang.String
     **/
    public static String getBaseDir(String regex){
        String baseDir=FileComments.class.getClassLoader().getResource(".").getPath();
        switch (regex.toLowerCase()){
            case "download":baseDir=baseDir+"\\download\\";break;
        }
        if(getOSName("Windows")){
            baseDir=baseDir.substring(1);
            baseDir=baseDir.replace("/target/classes/","\\src\\main\\resources\\static");
        }
        return baseDir;
    }

    /*
     * 输入的参数与当前系统的OSNAME比较
     * @Author wei
     * @Date 17:57 2020/12/17
     * @Param [osName]
     * @return boolean
     **/
    public static boolean getOSName(String osName){
        return System.getProperty("os.name").toUpperCase().startsWith(osName.toUpperCase());
    }



}
