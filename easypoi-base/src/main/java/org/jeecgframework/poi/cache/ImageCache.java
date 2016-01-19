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
package org.jeecgframework.poi.cache;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.apache.poi.util.IOUtils;
import org.jeecgframework.poi.cache.manager.POICacheManager;

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;

/**
 * 图片缓存处理
 * @author JueYue
 *  2016年1月8日 下午4:16:32
 */
public class ImageCache {

    private static LoadingCache<String, byte[]> loadingCache;

    static {
        loadingCache = CacheBuilder.newBuilder().expireAfterWrite(1, TimeUnit.DAYS)
            .maximumSize(2000).build(new CacheLoader<String, byte[]>() {
                @Override
                public byte[] load(String imagePath) throws Exception {
                    InputStream is = POICacheManager.getFile(imagePath);
                    BufferedImage bufferImg = ImageIO.read(is);
                    ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
                    try {
                        ImageIO.write(bufferImg,
                            imagePath.substring(imagePath.indexOf(".") + 1, imagePath.length()),
                            byteArrayOut);
                        return byteArrayOut.toByteArray();
                    } finally {
                        IOUtils.closeQuietly(is);
                        IOUtils.closeQuietly(byteArrayOut);
                    }
                }
            });
    }

    public static byte[] getImage(String imagePath) {
        try {
            return loadingCache.get(imagePath);
        } catch (Exception e) {
            return null;
        }

    }
}
