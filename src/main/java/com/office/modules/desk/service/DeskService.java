package com.office.modules.desk.service;

import com.office.common.utils.PageUtils;
import com.office.modules.desk.model.entity.Desk;
import java.util.Map;

/**
 * @Author wei
 * @Date 2021/1/15 16:24
 **/
public interface DeskService {

    boolean save(Desk desk);

    PageUtils listAll(Map<String,Object> param);

}
