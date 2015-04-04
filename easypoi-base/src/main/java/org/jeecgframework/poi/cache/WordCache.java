package org.jeecgframework.poi.cache;

import java.io.InputStream;

import org.jeecgframework.poi.cache.manager.POICacheManager;
import org.jeecgframework.poi.word.entity.MyXWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * word 缓存中心
 * 
 * @author JueYue
 * @date 2014年7月24日 下午10:54:31
 */
public class WordCache {

    private static final Logger LOGGER = LoggerFactory.getLogger(WordCache.class);

    public static MyXWPFDocument getXWPFDocumen(String url) {
        InputStream is = null;
        try {
            is = POICacheManager.getFile(url);
            MyXWPFDocument doc = new MyXWPFDocument(is);
            return doc;
        } catch (Exception e) {
            LOGGER.error(e.getMessage(),e);
        } finally {
            try {
                is.close();
            } catch (Exception e) {
                LOGGER.error(e.getMessage(),e);
            }
        }
        return null;
    }

}
