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
package cn.afterturn.easypoi.cache;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;

import javax.imageio.ImageIO;

import org.apache.poi.util.IOUtils;

import cn.afterturn.easypoi.cache.manager.POICacheManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 图片缓存处理
 *
 * @author JueYue
 *         2016年1月8日 下午4:16:32
 */
public class ImageCache {

    private static final Logger LOGGER = LoggerFactory
            .getLogger(ImageCache.class);

    public static byte[] getImage(String imagePath) {
        InputStream is = POICacheManager.getFile(imagePath);
        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        try {
            BufferedImage bufferImg = ImageIO.read(is);
            ImageIO.write(bufferImg,
                    imagePath.substring(imagePath.lastIndexOf(".") + 1, imagePath.length()),
                    byteArrayOut);
            return byteArrayOut.toByteArray();
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            return null;
        } finally {
            IOUtils.closeQuietly(is);
            IOUtils.closeQuietly(byteArrayOut);
        }

    }

}
