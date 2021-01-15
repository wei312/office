package com.office.modules.api.controller;

import cn.hutool.core.date.DateUtil;
import com.office.common.utils.FileComments;
import com.office.common.utils.PageUtils;
import com.office.common.utils.R;
import com.office.common.utils.ShiroUtils;
import com.office.common.utils.Tools;
import com.office.modules.api.model.entity.ApiList;
import com.office.modules.api.service.ApiListService;
import com.office.modules.export.utils.ExcelUtil;
import com.office.modules.export.utils.WordUtil;
import com.office.modules.sys.entity.SysUserEntity;
import com.office.modules.sys.utils.DownloadUtil;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.shiro.authz.annotation.RequiresPermissions;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

/**
 *
 * @Author wei
 * @Date 2020/12/15 17:04
 *
 **/
@RestController
@RequestMapping("/api")
public class ApiListController {

    @Autowired
    ApiListService apiListService;

    /**
     * 新增
     * @return
     */
    @PostMapping("/save")
    @RequiresPermissions("api:save")
    public R save(@RequestBody ApiList apiList){
        apiList.setApiId(Tools.getUUID());
        SysUserEntity userEntity= ShiroUtils.getUserEntity();
        apiList.setCreator(userEntity.getUsername());
        apiList.setCreateTime(DateUtil.now());
        apiList.setUpdater(userEntity.getUsername());
        apiList.setUpdateTime(DateUtil.now());
        return R.ok(String.valueOf(apiListService.save(apiList)));
    }

    /**
     * 修改
     * @param apiList
     * @return
     */
    @PostMapping("/update")
    @RequiresPermissions("api:update")
    public R update(@RequestBody ApiList apiList){
        apiList.setUpdater(ShiroUtils.getUserEntity().getUsername());
        apiList.setUpdateTime(DateUtil.now());
        return R.ok(String.valueOf(apiListService.updateByApiId(apiList)));
    }


    /**
     * 单个API信息
     */
    @ResponseBody
    @GetMapping("/info")
    @RequiresPermissions("api:info")
    public R info(@RequestParam Map<String,Object> params){
        ApiList apiList = apiListService.getByApiId(params.get("appId").toString());
        return R.ok().put("data",apiList);
    }
    /**
     * 列表
     */
    @GetMapping("/list")
    @RequiresPermissions("api:list")
    public R list(@RequestParam Map<String,Object> params){
        PageUtils page = apiListService.queryPage(params);
        return R.ok().put("page", page);
    }

    /**
     * 删除/批量删除
     * @param ids
     * @return
     */
    @PostMapping("/delete")
    @RequiresPermissions("api:delete")
    public R delete(@RequestBody String[] ids){
        return R.ok(String.valueOf(apiListService.removeByApiIds(Arrays.asList(ids))));
    }

    /**
     * 导出EXCEL
     * @return
     * @throws Exception
     */
    @PostMapping("/export/excel")
    @RequiresPermissions("api:export:excel")
    public R exportExcel(@RequestBody List<LinkedHashMap> lists )throws Exception{
        String fileName="接口清单.xls";
        String filePath= FileComments.getBaseDir("download")+fileName;
        String[] columns= new String[]{"apiParentModel","apiUrl","apiName","apiMethod"};
        String[] columnsName=new String[]{"所属模块","接口地址","接口名称","请求方式"};
        ExcelUtil.WriteToDocExcel(fileName,lists,columns,columnsName,filePath);
        return R.ok(fileName);
    }

    /**
     * 导出WORD
     * @param lists
     * @return
     * @throws Exception
     */
    @PostMapping("/export/word")
    @RequiresPermissions("api:export:word")
    public R excelWord(@RequestBody List<ApiList> lists) throws Exception {
        String fileName="接口调用清单.docx";
        String filePath= FileComments.getBaseDir("download")+fileName;
        String[] columns =new String[]{};
        LinkedHashMap params=WordUtil.getLinkedHashMap(lists,columns);
        WordUtil.docs2Word(params,filePath);
        return R.ok(fileName);
    }

    /**
     * 下载
     * @param fileName
     * @return
     * @throws Exception
     */
    @GetMapping("/download")
    public void downloadFiles(String fileName,HttpServletResponse response, HttpServletRequest request) throws Exception{
        String filePath= FileComments.getBaseDir("download")+fileName;
        DownloadUtil.downloadFile(response,filePath);
    }
}
