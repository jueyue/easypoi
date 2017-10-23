package cn.afterturn.easypoi.cache;

import java.util.concurrent.TimeUnit;

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;

import cn.afterturn.easypoi.excel.entity.ExcelToHtmlParams;
import cn.afterturn.easypoi.excel.html.ExcelToHtmlService;

/**
 * Excel 转变成为Html 的缓存
 * @author JueYue
 *  2015年8月7日 下午1:29:47
 */
public class HtmlCache {

    private static LoadingCache<ExcelToHtmlParams, String> loadingCache;

    static {
        loadingCache = CacheBuilder.newBuilder().expireAfterWrite(7, TimeUnit.DAYS).maximumSize(200)
            .build(new CacheLoader<ExcelToHtmlParams, String>() {
                @Override
                public String load(ExcelToHtmlParams params) throws Exception {
                    return new ExcelToHtmlService(params).printPage();
                }
            });
    }

    public static String getHtml(ExcelToHtmlParams params) {
        try {
            return loadingCache.get(params);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }
    
    public static void setLoadingCache(LoadingCache<ExcelToHtmlParams, String> loadingCache) {
        HtmlCache.loadingCache = loadingCache;
    }

}
