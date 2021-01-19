package com.office.modules.desk.controller;

import com.office.common.utils.PageUtils;
import com.office.common.utils.R;
import com.office.common.validator.ValidatorUtils;
import com.office.modules.app.entity.UserEntity;
import com.office.modules.app.form.RegisterForm;
import com.office.modules.desk.model.entity.Desk;
import com.office.modules.desk.service.DeskService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import java.util.Date;
import java.util.List;
import java.util.Map;
import org.apache.commons.codec.digest.DigestUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @Author wei
 * @Date 2021/1/15 16:22
 **/
@RestController
@RequestMapping("/desk")
@Api("DESK接口")
public class DeskController{

    @Autowired
    DeskService deskService;

    @PostMapping("/insert")
    @ApiOperation("新增")
    public R register(@RequestBody Desk desk){
        deskService.save(desk);
        return R.ok();
    }

    public R listAll(Map<String,Object> param){
        PageUtils page  =deskService.listAll(param);
        return R.ok().put("page",page);
    }

}
