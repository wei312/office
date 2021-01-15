package com.office.modules.export.utils;/**
 * Created by wei on 2020/12/17 16:55
 */


import java.lang.reflect.Field;

/**
 *
 * @Author wei
 * @Date 2020/12/17 16:55
 *
 **/
public class ClassUtil {

    /**
     * 反射获取字段列
     * @param clz
     * @return
     * @throws Exception
     */
    public static String[] getClassColumns(Class clz) throws Exception{
        Field[] fields = clz.getDeclaredFields();
        String[] string =new String[fields.length];
        for (int i = 0; i < fields.length; i++) {
            Field field=fields[i];
            string[i]=field.getName();
        }
        return string;
    }

}
