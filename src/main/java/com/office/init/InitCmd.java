package com.office.init;/**
 * Created by wei on 2020/12/16 16:29
 */


import org.springframework.beans.factory.InitializingBean;
import org.springframework.stereotype.Component;

/**
 *
 * @Author wei
 * @Date 2020/12/16 16:29
 *
 **/
//@Component
public class InitCmd implements InitializingBean {

    private static String cmdStart="cmd.exe /c start ";
    private static String cmdStop="cmd.exe /c start wmic process where name='cmd.exe' call terminate";
    private static String cmdTask="taskkill /f /t /im /consul.exe";
    @Override
    public void afterPropertiesSet() throws Exception {
        initCmdConsul("C:\\Users\\wei\\Desktop\\consul.bat");
    }

    public void initCmdConsul(String filePath) throws Exception {
        Process process = deleteCmd();
        process.waitFor();
        process.destroy();
        process=deleteCmdConsul();
        process.waitFor();
        process.destroy();

        Runtime.getRuntime().exec(cmdStart +filePath);
    }
    public Process deleteCmd()throws Exception{
        return Runtime.getRuntime().exec(cmdStop);
    }
    public Process deleteCmdConsul()throws Exception{
        return Runtime.getRuntime().exec(cmdTask);
    }

    //>taskkill /f /t /im /consul.exe
}
