package org.jeecgframework.poi.cache.manager;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;

/**
 * 缓存管理
 * 
 * @author JueYue
 * @date 2014年2月10日
 * @version 1.0
 */
public final class POICacheManager {

	private static LoadingCache<String, byte[]> loadingCache;

	static {
		loadingCache = CacheBuilder.newBuilder()
				.expireAfterWrite(1, TimeUnit.DAYS).maximumSize(50)
				.build(new CacheLoader<String, byte[]>() {
					@Override
					public byte[] load(String url) throws Exception {
						return new FileLoade().getFile(url);
					}
				});
	}
	public static InputStream getFile(String id) {
		try {
			//复杂数据,防止操作原数据
			byte[] result = Arrays.copyOf(loadingCache.get(id), loadingCache.get(id).length);
			return new ByteArrayInputStream(result);
		} catch (ExecutionException e) {
			e.printStackTrace();
		}
		return null;
	}

}
