package com.office.modules.desk.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.office.common.utils.PageUtils;
import com.office.common.utils.Query;
import com.office.modules.desk.mapper.DeskMapper;
import com.office.modules.desk.model.entity.Desk;
import com.office.modules.desk.service.DeskService;
import java.util.Map;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

/**
 * @Author wei
 * @Date 2021/1/15 16:24
 **/
@Service("deskService")
public class DeskServiceImpl extends ServiceImpl<DeskMapper,Desk> implements DeskService {

    @Autowired
    DeskMapper deskMapper;

    @Override
    public boolean save(Desk desk) {
        return deskMapper.save(desk);
    }

    @Override
    public PageUtils listAll(Map<String,Object> param){
        IPage<Desk> page = this.page(
                new Query<Desk>().getPage(param),
                new QueryWrapper<Desk>()
        );
        return new PageUtils(page);
    }
}
