package org.jeecgframework.poi.test.entity;

import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.annotation.ExcelVerify;

/**
 * Excel导入校验类
 * @author JueYue
 * @date 2015年2月24日 下午2:21:07
 */
public class ExcelVerifyEntity {

    /**
     * Email校验
     */
    @Excel(name = "Email", width = 25)
    @ExcelVerify(isEmail = true, notNull = true)
    private String email;
    /**
     * 手机号校验
     */
    @Excel(name = "Mobile", width = 20)
    @ExcelVerify(isMobile = true, notNull = true)
    private String mobile;
    /**
     * 电话校验
     */
    @Excel(name = "Tel", width = 20)
    @ExcelVerify(isTel = true, notNull = true)
    private String tel;
    /**
     * 最长校验
     */
    @Excel(name = "MaxLength")
    @ExcelVerify(maxLength = 15)
    private String maxLength;
    /**
     * 最短校验
     */
    @Excel(name = "MinLength")
    @ExcelVerify(minLength = 3)
    private String minLength;
    /**
     * 非空校验
     */
    @Excel(name = "NotNull")
    @ExcelVerify(notNull = true)
    private String notNull;
    /**
     * 正则校验
     */
    @Excel(name = "Regex")
    @ExcelVerify(regex = "[\u4E00-\u9FA5]*", regexTip = "不是中文")
    private String regex;
    /**
     * 接口校验
     */
    private String interHandler;

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getMobile() {
        return mobile;
    }

    public void setMobile(String mobile) {
        this.mobile = mobile;
    }

    public String getTel() {
        return tel;
    }

    public void setTel(String tel) {
        this.tel = tel;
    }

    public String getMaxLength() {
        return maxLength;
    }

    public void setMaxLength(String maxLength) {
        this.maxLength = maxLength;
    }

    public String getMinLength() {
        return minLength;
    }

    public void setMinLength(String minLength) {
        this.minLength = minLength;
    }

    public String getNotNull() {
        return notNull;
    }

    public void setNotNull(String notNull) {
        this.notNull = notNull;
    }

    public String getRegex() {
        return regex;
    }

    public void setRegex(String regex) {
        this.regex = regex;
    }

    public String getInterHandler() {
        return interHandler;
    }

    public void setInterHandler(String interHandler) {
        this.interHandler = interHandler;
    }

}
