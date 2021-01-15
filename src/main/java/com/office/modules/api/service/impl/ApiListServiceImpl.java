package com.office.modules.api.service.impl;/**
 * Created by wei on 2020/12/15 17:12
 */

import com.baomidou.mybatisplus.annotation.TableName;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.conditions.update.UpdateWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.office.common.utils.PageUtils;
import com.office.common.utils.Query;
import com.office.modules.api.mapper.ApiListMapper;
import com.office.modules.api.model.entity.ApiList;
import com.office.modules.api.model.param.ApiListParam;
import com.office.modules.api.service.ApiListService;
import com.office.modules.sys.dao.SysLogDao;
import com.office.modules.sys.entity.SysLogEntity;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import org.apache.commons.lang.StringUtils;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

/**
 *
 * @Author wei
 * @Date 2020/12/15 17:12
 *
 **/
@Service("apiListService")
public class ApiListServiceImpl extends ServiceImpl<ApiListMapper, ApiList>  implements ApiListService {

    @Override
    public ApiList getByApiId(String id){
        return this.getOne(new QueryWrapper<ApiList>().lambda().eq(ApiList::getApiId,id));
    }

    @Override
    public boolean updateByApiId(ApiList apiList){
        return this.update(apiList,new UpdateWrapper<ApiList>().lambda().eq(ApiList::getApiId,apiList.getApiId()));
    }

    @Override
    public PageUtils queryPage(Map<String,Object> param) {
        String apiName = (String)param.get("apiName");

        IPage<ApiList> page = this.page(
                new Query<ApiList>().getPage(param),
                new QueryWrapper<ApiList>().lambda().like(StringUtils.isNotEmpty(apiName),ApiList::getApiName, apiName)
                .orderBy(true,false,ApiList::getCreateTime)
        );

        return new PageUtils(page);
    }

    public boolean removeByApiIds(List<String> apiIds){
        return this.remove(new QueryWrapper<ApiList>().lambda().in(ApiList::getApiId, apiIds));
    }
}
