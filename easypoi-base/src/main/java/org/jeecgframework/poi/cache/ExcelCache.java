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

import java.io.InputStream;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jeecgframework.poi.cache.manager.POICacheManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel类型的缓存
 * 
 * @author JueYue
 *  2014年2月11日
 * @version 1.0
 */
public final class ExcelCache {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelCache.class);

    public static Workbook getWorkbook(String url, Integer[] sheetNums, boolean needAll) {
        InputStream is = null;
        List<Integer> sheetList = Arrays.asList(sheetNums);
        try {
            is = POICacheManager.getFile(url);
            Workbook wb = WorkbookFactory.create(is);
            // 删除其他的sheet
            if (!needAll) {
                for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--) {
                    if (!sheetList.contains(i)) {
                        wb.removeSheetAt(i);
                    }
                }
            }
            return wb;
        } catch (InvalidFormatException e) {
            LOGGER.error(e.getMessage(), e);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        } finally {
            try {
                is.close();
            } catch (Exception e) {
                LOGGER.error(e.getMessage(), e);
            }
        }
        return null;
    }

}
