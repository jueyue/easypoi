package org.jeecgframework.poi.cache;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jeecgframework.poi.cache.manager.POICacheManager;

/**
 * Excel类型的缓存
 * 
 * @author JueYue
 * @date 2014年2月11日
 * @version 1.0
 */
public final class ExcelCache {

	public static Workbook getWorkbook(String url, int index) {
		InputStream is = null;
		try {
			is = POICacheManager.getFile(url);
			Workbook wb = WorkbookFactory.create(is);
			// 删除其他的sheet
			for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--) {
				if (i != index) {
					wb.removeSheetAt(i);
				}
			}
			return wb;
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return null;
	}

}
