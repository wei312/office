package com.office.modules.sys.utils;/**
 * Created by wei on 2020/12/24 11:19
 */

import com.office.modules.sys.oauth2.OAuth2Token;
import javax.servlet.http.HttpServletRequest;
import org.apache.commons.lang.StringUtils;

/**
 *
 * @Author wei
 * @Date 2020/12/24 11:19
 *
 **/
public class TokenUtil {

    static HttpServletRequest httpRequest;
    /**
     * 获取请求的token
     */
    public static String getRequestToken(){
        return new OAuth2Token("admin").getPrincipal();
    }

}
