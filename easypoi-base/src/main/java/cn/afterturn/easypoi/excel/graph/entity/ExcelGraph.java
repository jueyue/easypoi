/**
 * 
 */
package cn.afterturn.easypoi.excel.graph.entity;

import java.util.List;

/**
 * @author xfworld
 * @since 2016-1-7
 * @version 1.0    
 * @see com.dawnpro.core.export.excel.model.ExcelGraph
 * 
 */
public interface ExcelGraph
{
	public ExcelGraphElement getCategory();
	public List<ExcelGraphElement> getValueList();
	public Integer getGraphType();
	public List<ExcelTitleCell> getTitleCell();
	public List<String> getTitle();
}
