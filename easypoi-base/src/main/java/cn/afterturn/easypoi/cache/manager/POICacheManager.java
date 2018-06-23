/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cn.afterturn.easypoi.cache.manager;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Arrays;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 缓存管理
 *
 * @author JueYue
 *  2014年2月10日
 *  2015年10月17日
 * @version 1.1
 */
public final class POICacheManager {

    private static final Logger LOGGER = LoggerFactory
            .getLogger(POICacheManager.class);

    private static IFileLoader fileLoader = new FileLoaderImpl();

    private static ThreadLocal<IFileLoader> LOCAL_FILELOADER = new ThreadLocal<IFileLoader>();

    public static InputStream getFile(String id) {
        try {
            byte[] result;
            //复杂数据,防止操作原数据
            if (LOCAL_FILELOADER.get() != null) {
                result = LOCAL_FILELOADER.get().getFile(id);
            }
            result = fileLoader.getFile(id);
            result = Arrays.copyOf(result, result.length);
            return new ByteArrayInputStream(result);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
        return null;
    }

    public static void setFileLoader(IFileLoader fileLoader) {
        POICacheManager.fileLoader = fileLoader;
    }

    /**
     * 一次线程有效
     * @param fileLoader
     */
    public static void setFileLoaderOnce(IFileLoader fileLoader) {
        if (fileLoader != null) {
            LOCAL_FILELOADER.set(fileLoader);
        }
    }

}
