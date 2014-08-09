package org.jeecgframework.poi.word.parse.excel;

import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jeecgframework.poi.word.util.ParseWordUtil;

/**
 * 处理和生成Map 类型的数据变成表格
 * @author JueYue
 * @date 2014年8月9日 下午10:28:46
 */
public final class ExcelMapParse {
	
	/**
	 * 解析下一行,并且生成更多的行
	 * 
	 * @Author JueYue
	 * @date 2013-11-18
	 * @param table
	 * @param listobj2
	 */
	public static void parseNextRowAndAddRow(XWPFTable table, int index,
			List<Object> list) throws Exception {
		XWPFTableRow currentRow = table.getRow(index);
		String[] params = parseCurrentRowGetParams(currentRow);
		table.removeRow(index);// 移除这一行
		int cellIndex = 0;// 创建完成对象一行好像多了一个cell
		for (Object obj : list) {
			currentRow = table.createRow();
			for (cellIndex = 0; cellIndex < currentRow.getTableCells().size(); cellIndex++) {
				currentRow
						.getTableCells()
						.get(cellIndex)
						.setText(
								ParseWordUtil.getValueDoWhile(obj,
										params[cellIndex].split("\\."), 0)
										.toString());
			}
			for (; cellIndex < params.length; cellIndex++) {
				currentRow.createCell().setText(
						ParseWordUtil.getValueDoWhile(obj,
								params[cellIndex].split("\\."), 0).toString());
			}
		}

	}
	
	/**
	 * 解析参数行,获取参数列表
	 * 
	 * @Author JueYue
	 * @date 2013-11-18
	 * @param currentRow
	 * @return
	 */
	private static String[] parseCurrentRowGetParams(XWPFTableRow currentRow) {
		List<XWPFTableCell> cells = currentRow.getTableCells();
		String[] params = new String[cells.size()];
		String text;
		for (int i = 0; i < cells.size(); i++) {
			text = cells.get(i).getText();
			params[i] = text == null ? "" : text.trim().replace("{{", "")
					.replace("}}", "");
		}
		return params;
	}

}
