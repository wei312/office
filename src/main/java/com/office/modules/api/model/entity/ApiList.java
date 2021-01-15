package com.office.modules.api.model.entity;/**
 * Created by wei on 2020/12/15 17:08
 */

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import io.lettuce.core.dynamic.annotation.Key;
import java.io.Serializable;
import lombok.Data;

/**
 *
 * @Author wei
 * @Date 2020/12/15 17:08
 *
 **/
@TableName("api_list")
@Data
public class ApiList implements Serializable {
    @TableId(value="api_id",type = IdType.ASSIGN_ID)
    private String apiId;
    private String apiUrl;
    private String apiName;
    private String apiMethod;
    private String apiRequestType;
    private String apiResponseType;
    private String apiParentModel;
    private String apiVersion;
    private String example;
    private String requestExample;
    private String responseExample;
    private String creator;
    private String createTime;
    private String updater;
    private String updateTime;
}
