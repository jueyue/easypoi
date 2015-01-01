package org.jeecgframework.poi.excel.imports.sax.parse;

import java.util.List;

import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;

public interface ISaxRowRead {
    /**
     * 获取返回数据
     * @param <T>
     * @return
     */
    public <T> List<T> getList();
    /**
     * 解析数据
     * @param index
     * @param datas
     */
    public void parse(int index, List<SaxReadCellEntity> datas);


}
