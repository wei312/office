package com.office.common.utils;/**
 * Created by wei on 2020/12/16 14:04
 */

import java.util.UUID;

/**
 *
 * @Author wei
 * @Date 2020/12/16 14:04
 *
 **/
public class Tools {

    public static String getUUID() {
        return UUID.randomUUID().toString().replace("-","");
    }


}
