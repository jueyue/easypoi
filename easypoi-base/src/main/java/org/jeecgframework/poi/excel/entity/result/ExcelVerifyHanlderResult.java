package org.jeecgframework.poi.excel.entity.result;

/**
 * Excel导入处理返回结果
 * 
 * @author JueYue
 * @date 2014年6月23日 下午11:03:29
 */
public class ExcelVerifyHanlderResult {
    /**
     * 是否正确
     */
    private boolean success;
    /**
     * 错误信息
     */
    private String  msg;

    public ExcelVerifyHanlderResult() {

    }

    public ExcelVerifyHanlderResult(boolean success) {
        this.success = success;
    }

    public ExcelVerifyHanlderResult(boolean success, String msg) {
        this.success = success;
        this.msg = msg;
    }

    public String getMsg() {
        return msg;
    }

    public boolean isSuccess() {
        return success;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }

}
