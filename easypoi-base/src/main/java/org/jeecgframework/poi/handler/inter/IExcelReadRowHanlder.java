package org.jeecgframework.poi.handler.inter;

/**
 * 接口自定义处理类
 * @author JueYue
 * @date 2015年1月16日 下午8:06:26
 * @param <T>
 */
public interface IExcelReadRowHanlder<T> {
    /**
     * 处理解析对象
     * @param t
     */
    public void hanlder(T t);

}
