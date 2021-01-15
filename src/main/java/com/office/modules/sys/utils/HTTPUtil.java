package com.office.modules.sys.utils;/**
 * Created by wei on 2020/12/24 14:34
 */

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

/**
 *
 * @Author wei
 * @Date 2020/12/24 14:34
 *
 **/
public class HTTPUtil {

    public static ServletRequestAttributes getServletRequestAttributes(){
        ServletRequestAttributes servletRequestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        return servletRequestAttributes;
    }

    public static  HttpServletRequest getRequest(){
        HttpServletRequest request = getServletRequestAttributes().getRequest();
        return request;
    }

    public static HttpServletResponse getResponse(){
        HttpServletResponse response = getServletRequestAttributes().getResponse();
        return response;
    }

}
