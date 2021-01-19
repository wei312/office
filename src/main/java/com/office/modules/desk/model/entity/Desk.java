package com.office.modules.desk.model.entity;

import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;

/**
 * @Author wei
 * @Date 2021/1/15 16:14
 **/
@TableName("desk")
@Data
@NoArgsConstructor
@EqualsAndHashCode(callSuper = false)
public class Desk {

    private String id;
    private String parentId;
    private String name;
    private String path;
    private long size;
    private String type;
    private String fileLastTime;
    private String createTime;
    


}
