package com.office.modules.desk.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.office.modules.desk.model.entity.Desk;
import org.apache.ibatis.annotations.Mapper;

/**
 * @Author wei
 * @Date 2021/1/15 16:26
 **/
@Mapper
public interface DeskMapper extends BaseMapper<Desk> {

    boolean save(Desk desk);

}
