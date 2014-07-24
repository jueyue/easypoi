package org.jeecgframework.poi.excel.annotation;

/**
 * Excel 导入校验
 * 
 * @author JueYue
 * @date 2014年6月23日 下午10:46:26
 */
public @interface ExcelVerify {
	/**
	 * 接口校验
	 * @return
	 */
	public boolean interHandler() default true;

	/**
	 * 不允许空
	 * 
	 * @return
	 */
	public boolean notNull() default false;

	/**
	 * 是13位移动电话
	 * 
	 * @return
	 */
	public boolean isMobile() default false;
	/**
	 * 是座机号码
	 * 
	 * @return
	 */
	public boolean isTel() default false;

	/**
	 * 是电子邮件
	 * 
	 * @return
	 */
	public boolean isEmail() default false;

	/**
	 * 最小长度
	 * 
	 * @return
	 */
	public int minLength() default -1;

	/**
	 * 最大长度
	 * 
	 * @return
	 */
	public int maxLength() default -1;

	/**
	 * 正在表达式
	 * 
	 * @return
	 */
	public String regex() default "";
	/**
	 * 正在表达式,错误提示信息
	 * 
	 * @return
	 */
	public String regexTip() default "数据不符合规范";

}
