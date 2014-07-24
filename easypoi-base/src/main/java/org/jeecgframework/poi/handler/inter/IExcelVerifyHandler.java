package org.jeecgframework.poi.handler.inter;

import org.jeecgframework.poi.excel.entity.result.ExcelVerifyHanlderResult;
/**
 * 导入校验接口
 * @author JueYue
 * @date 2014年6月23日 下午11:08:21
 */
public interface IExcelVerifyHandler {
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
	public ExcelVerifyHanlderResult verifyHandler(Object obj, String name,
			Object value);

}
