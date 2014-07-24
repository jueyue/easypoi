package org.jeecgframework.poi.cache;

import java.io.IOException;
import java.io.InputStream;

import org.jeecgframework.poi.cache.manager.POICacheManager;
import org.jeecgframework.poi.word.entity.JeecgXWPFDocument;

/**
 * word 缓存中心
 * 
 * @author JueYue
 * @date 2014年7月24日 下午10:54:31
 */
public class WordCache {

	public static JeecgXWPFDocument getXWPFDocumen(String url) {
		InputStream is = null;
		try {
			is = POICacheManager.getFile(url);
			JeecgXWPFDocument doc = new JeecgXWPFDocument(is);
			return doc;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return null;
	}

}
