/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.jeecgframework.poi.cache.manager;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;

/**
 * 缓存管理
 * 
 * @author JueYue
 *  2014年2月10日
 *  2015年10月17日
 * @version 1.1
 */
public final class POICacheManager {

    private static final Logger                 LOGGER           = LoggerFactory
        .getLogger(POICacheManager.class);

    private static LoadingCache<String, byte[]> loadingCache;

    private static IFileLoader                  fileLoder;

    private static ThreadLocal<IFileLoader>     LOCAL_FILELOADER = new ThreadLocal<IFileLoader>();

    static {
        loadingCache = CacheBuilder.newBuilder().expireAfterWrite(1, TimeUnit.HOURS).maximumSize(50)
            .build(new CacheLoader<String, byte[]>() {
                @Override
                public byte[] load(String url) throws Exception {
                    if (LOCAL_FILELOADER.get() != null)
                        return LOCAL_FILELOADER.get().getFile(url);
                    return fileLoder.getFile(url);
                }
            });
        //设置默认实现
        fileLoder = new FileLoadeImpl();
    }

    public static InputStream getFile(String id) {
        try {
            //复杂数据,防止操作原数据
            byte[] result = Arrays.copyOf(loadingCache.get(id), loadingCache.get(id).length);
            return new ByteArrayInputStream(result);
        } catch (ExecutionException e) {
            LOGGER.error(e.getMessage(), e);
        }
        return null;
    }

    public static void setFileLoder(IFileLoader fileLoder) {
        POICacheManager.fileLoder = fileLoder;
    }

    /**
     * 一次线程有效
     * @param fileLoder
     */
    public static void setFileLoderOnce(IFileLoader fileLoder) {
        if (fileLoder != null)
            LOCAL_FILELOADER.set(fileLoder);
    }

    public static void setLoadingCache(LoadingCache<String, byte[]> loadingCache) {
        POICacheManager.loadingCache = loadingCache;
    }

}
