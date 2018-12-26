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
package cn.afterturn.easypoi.handler.inter;

/**
 * 接口自定义处理类
 * @author JueYue
 *  2015年1月16日 下午8:06:26
 * @param <T>
 */
public interface IReadHandler<T> {
    /**
     * 处理解析对象
     * @param t
     */
    public void handler(T t);


    /**
     * 处理完成之后的业务
     */
    public void doAfterAll();

}
