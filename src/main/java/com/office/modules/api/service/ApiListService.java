package com.office.modules.api.service;/**
 * Created by wei on 2020/12/15 17:07
 */

import com.baomidou.mybatisplus.extension.service.IService;
import com.office.common.utils.PageUtils;
import com.office.modules.api.model.entity.ApiList;
import java.util.List;
import java.util.Map;

/**
 *
 * @Author wei
 * @Date 2020/12/15 17:07
 *
 **/
public interface ApiListService extends IService<ApiList> {

    /*
     * 查询API BY ID
     * @param id
     * @return com.office.modules.api.model.entity.ApiList
     **/
    ApiList getByApiId(String id);

    boolean updateByApiId(ApiList apiList);

    PageUtils queryPage( Map<String,Object> param);

    boolean removeByApiIds(List<String> ApiIds);

}
