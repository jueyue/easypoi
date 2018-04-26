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

import cn.afterturn.easypoi.util.PoiPublicUtil;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.net.URL;
import java.net.URLConnection;

/**
 * 文件加载类,根据路径加载指定文件
 * @author JueYue
 *  2014年2月10日
 * @version 1.0
 */
public class FileLoaderImpl implements IFileLoader {

    private static final Logger LOGGER = LoggerFactory.getLogger(FileLoaderImpl.class);

    @Override
    public byte[] getFile(String url) {
        InputStream fileis = null;
        ByteArrayOutputStream baos = null;
        try {

            //判断是否是网络地址
            if (url.startsWith("http")) {
                URL urlObj = new URL(url);
                URLConnection urlConnection = urlObj.openConnection();
                urlConnection.setConnectTimeout(30);
                urlConnection.setReadTimeout(60);
                urlConnection.setDoInput(true);
                fileis = urlConnection.getInputStream();
            } else {
                //先用绝对路径查询,再查询相对路径
                try {
                    fileis = new FileInputStream(url);
                } catch (FileNotFoundException e) {
                    //获取项目文件
                    fileis = FileLoaderImpl.class.getClassLoader().getResourceAsStream(url);
                }
            }

            baos = new ByteArrayOutputStream();
            byte[] buffer = new byte[1024];
            int len;
            while ((len = fileis.read(buffer)) > -1) {
                baos.write(buffer, 0, len);
            }
            baos.flush();
            return baos.toByteArray();
        } catch (IOException e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            IOUtils.closeQuietly(fileis);
            IOUtils.closeQuietly(baos);
        }
        LOGGER.error(fileis + "这个路径文件没有找到,请查询");
        return null;
    }

}
