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
package org.jeecgframework.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 导出基本注释
 * @author JueYue
 *  2014年6月20日 下午10:25:12
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Excel {

    /**
     * 导出时间设置,如果字段是Date类型则不需要设置 数据库如果是string 类型,这个需要设置这个数据库格式
     */
    public String databaseFormat() default "yyyyMMddHHmmss";

    /**
     * 导出的时间格式,以这个是否为空来判断是否需要格式化日期
     */
    public String exportFormat() default "";

    /**
     * 时间格式,相当于同时设置了exportFormat 和 importFormat
     */
    public String format() default "";

    /**
     * 导出时在excel中每个列的高度 单位为字符，一个汉字=2个字符
     */
    public double height() default 10;

    /**
     * 导出类型 1 从file读取 2 是从数据库中读取 默认是文件 同样导入也是一样的
     */
    public int imageType() default 1;

    /**
     * 导入的时间格式,以这个是否为空来判断是否需要格式化日期
     */
    public String importFormat() default "";

    /**
     * 文字后缀,如% 90 变成90%
     */
    public String suffix() default "";

    /**
     * 是否换行 即支持\n
     */
    public boolean isWrap() default true;

    /**
     * 合并单元格依赖关系,比如第二列合并是基于第一列 则{1}就可以了
     */
    public int[] mergeRely() default {};

    /**
     * 纵向合并内容相同的单元格
     */
    public boolean mergeVertical() default false;

    /**
     * 导出时，对应数据库的字段 主要是用户区分每个字段， 不能有annocation重名的 导出时的列名
     * 导出排序跟定义了annotation的字段的顺序有关 可以使用a_id,b_id来确实是否使用
     */
    public String name();

    /**
     * 是否需要纵向合并单元格(用于含有list中,单个的单元格,合并list创建的多个row)
     */
    public boolean needMerge() default false;

    /**
     * 展示到第几个可以使用a_id,b_id来确定不同排序
     */
    public String orderNum() default "0";

    /**
     * 值得替换  导出是{a_id,b_id} 导入反过来,所以只用写一个
     */
    public String[] replace() default {};

    /**
     * 导入路径,如果是图片可以填写,默认是upload/className/ IconEntity这个类对应的就是upload/Icon/
     * 
     */
    public String savePath() default "upload";

    /**
     * 导出类型 1 是文本 2 是图片,3 是函数,10 是数字 默认是文本
     */
    public int type() default 1;

    /**
     * 导出时在excel中每个列的宽 单位为字符，一个汉字=2个字符 如 以列名列内容中较合适的长度 例如姓名列6 【姓名一般三个字】
     * 性别列4【男女占1，但是列标题两个汉字】 限制1-255
     */
    public double width() default 10;

    /**
     * 是否自动统计数据,如果是统计,true的话在最后追加一行统计,把所有数据都和
     * 这个处理会吞没异常,请注意这一点
     * @return
     */
    public boolean isStatistics() default false;
    
    /**
     * 这个是不是超链接,如果是需要实现接口返回对象
     * @return
     */
    public boolean isHyperlink() default false;
    
    /**
     *  导入时会校验这个字段,看看这个字段是不是导入的Excel中有,如果没有说明是错误的Excel
     *  本意是想用true的,想想还是false比较好
     *  可以使用a_id,b_id来确实是否使用
     * @return
     */
    public String isImportField() default "false";
}
