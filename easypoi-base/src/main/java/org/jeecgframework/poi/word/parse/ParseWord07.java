package org.jeecgframework.poi.word.parse;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jeecgframework.poi.cache.WordCache;
import org.jeecgframework.poi.word.entity.JeecgXWPFDocument;
import org.jeecgframework.poi.word.entity.WordImageEntity;
import org.jeecgframework.poi.word.util.ParseWordUtil;

/**
 * 解析07版的Word,替换文字,生成表格,生成图片
 * 
 * @author JueYue
 * @date 2013-11-16
 * @version 1.0
 */
@SuppressWarnings({"unchecked","rawtypes"})
public class ParseWord07 {

	/**
	 * 解析07版的Word并且进行赋值
	 * 
	 * @Author JueYue
	 * @date 2013-11-16
	 * @return
	 * @throws Exception
	 */
	public  XWPFDocument parseWord(String url, Map<String, Object> map)
			throws Exception {
		JeecgXWPFDocument doc = WordCache.getXWPFDocumen(url);
		parseWordSetValue(doc, map);
		return doc;
	}

	private  void parseWordSetValue(JeecgXWPFDocument doc,
			Map<String, Object> map) throws Exception {
		//第一步解析文档
		parseAllParagraphic(doc.getParagraphs(),map);
		//第二步解析所有表格
		XWPFTable table;
		Iterator<XWPFTable> itTable = doc.getTablesIterator();
		while (itTable.hasNext()) {
			table = itTable.next();
			if (table.getText().indexOf("{{") != -1) {
				parseThisTable(table, map);
			}
		}
		
	}

	/**
	 * 解析所有的文本
	 *@Author JueYue
	 *@date   2013-11-17
	 *@param paragraphs
	 *@param map
	 */
	private  void parseAllParagraphic(List<XWPFParagraph> paragraphs,
			Map<String, Object> map) throws Exception  {
		XWPFParagraph paragraph;
		for (int i = 0; i < paragraphs.size(); i++) {
			paragraph = paragraphs.get(i);
			if (paragraph.getText().indexOf("{{") != -1) {
				parseThisParagraph(paragraph, map);
			}

		}
		
	}

	/**
	 * 解析这个表格
	 * 
	 * @Author JueYue
	 * @date 2013-11-17
	 * @param table
	 * @param map
	 */
	private  void parseThisTable(XWPFTable table, Map<String, Object> map)  throws Exception {
		XWPFTableRow row;
		List<XWPFTableCell> cells;
		Object listobj;
        for (int i = 0; i < table.getNumberOfRows(); i++) {
            row = table.getRow(i);
            cells = row.getTableCells();
            if(cells.size() == 1){
            	listobj = checkThisTableIsNeedIterator(cells.get(0),map);
            	if(listobj==null){
            		parseThisRow(cells,map);
            	}else{
            		table.removeRow(i);//删除这一行
            		parseNextRowAndAddRow(table,i,(List) listobj);
            	}
            }else{
            	parseThisRow(cells,map);
            }
        }
	}
	
	
	/**
	 * 解析下一行,并且生成更多的行
	 *@Author JueYue
	 *@date   2013-11-18
	 *@param table
	 *@param listobj2
	 */
	private  void parseNextRowAndAddRow(XWPFTable table,int index, List<Object> list) throws Exception  {
		XWPFTableRow currentRow = table.getRow(index);
		String[] params = parseCurrentRowGetParams(currentRow);
		table.removeRow(index);//移除这一行
		int cellIndex = 0;//创建完成对象一行好像多了一个cell
		for(Object obj :list){
			currentRow = table.createRow();
			for(cellIndex = 0;cellIndex<currentRow.getTableCells().size();cellIndex++){
				currentRow.getTableCells().get(cellIndex).setText(
						ParseWordUtil.getValueDoWhile(obj, params[cellIndex].split("\\."), 0).toString());
			}
			for (;cellIndex<params.length;cellIndex++) {
				currentRow.createCell().setText(
						ParseWordUtil.getValueDoWhile(obj, params[cellIndex].split("\\."), 0).toString());
			}	
		}
		
	}

	/**
	 * 解析参数行,获取参数列表
	 *@Author JueYue
	 *@date   2013-11-18
	 *@param currentRow
	 *@return
	 */
	private  String[] parseCurrentRowGetParams(XWPFTableRow currentRow) {
		List<XWPFTableCell> cells = currentRow.getTableCells();
		String[] params = new String[cells.size()];
		String text;
		for (int i=0;i<cells.size();i++) {
			text = cells.get(i).getText();
			params[i] = text==null?"":text.trim().replace("{{","").replace("}}","");
		}		
		return params;
	}

	private  void parseThisRow(List<XWPFTableCell> cells,
			Map<String, Object> map) throws Exception {
		for (XWPFTableCell cell : cells) {
			parseAllParagraphic(cell.getParagraphs(),map);
		}		
	}

	/**
	 *判断是不是迭代输出
	 *@Author JueYue
	 *@date   2013-11-18
	 *@return
	 * @throws Exception 
	 */
	private  Object checkThisTableIsNeedIterator(XWPFTableCell cell
			, Map<String, Object> map) throws Exception {
		String text = cell.getText().trim();
		//判断是不是迭代输出
		if(text.startsWith("{{")&&text.endsWith("}}")&&text.indexOf("in ")!=-1){
			return ParseWordUtil.getRealValue(text.replace("in ", "").trim(),map);
		}
		return null;
	}

	/**
	 * 解析这个段落
	 * 
	 * @Author JueYue
	 * @date 2013-11-16
	 * @param paragraph
	 * @param map
	 */
	private  void parseThisParagraph(XWPFParagraph paragraph,
			Map<String, Object> map) throws Exception  {
		XWPFRun run;
		XWPFRun currentRun = null;// 拿到的第一个run,用来set值,可以保存格式
		String currentText = "";// 存放当前的text
		String text;
		Boolean isfinde = false;// 判断是不是已经遇到{{
		List<Integer> runIndex = new ArrayList<Integer>();// 存储遇到的run,把他们置空
		for (int i = 0; i < paragraph.getRuns().size(); i++) {
			run = paragraph.getRuns().get(i);
			text = run.getText(0);
			if (text == null || text == "") {
				continue;
			}// 如果为空或者""这种这继续循环跳过
			if (isfinde) {
				currentText += text;
				if (currentText.indexOf("{{") == -1) {
					isfinde = false;
					runIndex.clear();
				} else {
					runIndex.add(i);
				}
				if (currentText.indexOf("}}") != -1) {
					changeValues(paragraph, currentRun, currentText, runIndex,
							map);
					currentText = "";
					isfinde = false;
				}
			} else if (text.indexOf("{") >= 0) {// 判断是不是开始
				currentText = text;
				isfinde = true;
				currentRun = run;
			}else{
				currentText = "";
			}
			if (currentText.indexOf("}}") != -1) {
				changeValues(paragraph, currentRun, currentText, runIndex, map);
				isfinde = false;
			}
		}

	}

	/**
	 * 根据条件改变值
	 * 
	 * @param map
	 * @Author JueYue
	 * @date 2013-11-16
	 */
	private  void changeValues(XWPFParagraph paragraph,
			XWPFRun currentRun, String currentText, List<Integer> runIndex,
			Map<String, Object> map) throws Exception  {
		Object obj = ParseWordUtil.getRealValue(currentText, map);
		if(obj instanceof WordImageEntity){//如果是图片就设置为图片
			currentRun.setText("", 0);
			addAnImage((WordImageEntity)obj,currentRun);
		}else{
			currentText = obj.toString();
			currentRun.setText(currentText,0);
		}
		for (int k = 0; k < runIndex.size(); k++) {
			paragraph.getRuns().get(runIndex.get(k)).setText("", 0);
		}
		runIndex.clear();
	}

	/**
	 * 添加图片
	 *@Author JueYue
	 *@date   2013-11-20
	 *@param obj
	 *@param currentRun
	 * @throws Exception 
	 */
	private void addAnImage(WordImageEntity obj, XWPFRun currentRun) throws Exception {
		Object [] isAndType = ParseWordUtil.getIsAndType(obj);
		String picId;
		try {
			picId = currentRun.getParagraph().getDocument().addPictureData(
					(byte[]) isAndType[0],
					(Integer) isAndType[1]);
			((JeecgXWPFDocument) currentRun.getParagraph().getDocument()).createPicture(
					currentRun,picId, 
					currentRun.getParagraph().getDocument().getNextPicNameNumber((Integer) isAndType[1]),
					obj.getWidth(),obj.getHeight());  
			
		} catch (Exception e) {
			e.printStackTrace();
		} 
		
	}

	
	

}
