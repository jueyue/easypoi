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

/**
 * 缓存读取
 * @author JueYue
 *         默认实现是FileLoader
 *  2015年10月17日 下午7:12:01
 */
public interface IFileLoader {
    /**
     * 可以自定义KEY的作用
     * @param key
     * @return
     */
    public byte[] getFile(String key);

}
