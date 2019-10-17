/**
 * 
 */
package cn.afterturn.easypoi.excel.entity;

import java.util.List;
import java.util.Map;

import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import lombok.Data;

/**
 * @author xfworld
 * @since 2016-5-26
 * @version 1.0
 */
@Data
public class ExportExcelItem {
	private List<ExcelExportEntity> entityList;
	private List<Map<String, Object>> resultList;
}
