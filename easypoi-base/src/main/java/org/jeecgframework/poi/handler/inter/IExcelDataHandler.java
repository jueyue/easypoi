package org.jeecgframework.poi.handler.inter;

/**
 * Excel 导入导出 数据处理接口
 * 
 * @author JueYue
 * @date 2014年6月19日 下午11:59:45
 */
public interface IExcelDataHandler {

	/**
	 * 获取需要处理的字段,导入和导出统一处理了, 减少书写的字段
	 * 
	 * @return
	 */
	public String[] getNeedHandlerFields();

	public void setNeedHandlerFields(String[] fields);

	/**
	 * 导出处理方法
	 * 
	 * @param obj
	 *            当前对象
	 * @param name
	 *            当前字段名称
	 * @param value
	 *            当前值
	 * @return
	 */
	public Object exportHandler(Object obj, String name, Object value);

	/**
	 * 导入处理方法 当前对象,当前字段名称,当前值
	 * 
	 * @param obj
	 *            当前对象
	 * @param name
	 *            当前字段名称
	 * @param value
	 *            当前值
	 * @return
	 */
	public Object importHandler(Object obj, String name, Object value);

}
