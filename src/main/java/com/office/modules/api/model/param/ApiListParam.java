package com.office.modules.api.model.param;/**
 * Created by wei on 2020/12/15 17:10
 */

import lombok.Data;

/**
 *
 * @Author wei
 * @Date 2020/12/15 17:10
 *
 **/
@Data
public class ApiListParam {

    private String apiId;
    private String apiUrl;
    private String apiName;
    private String apiMethod;
    private String apiParentModel;
    private String apiVersion;
    private String creator;
    private String createTime;
    private String updater;
    private String updateTime;
}
