package org.jeecgframework.poi.excel.imports.sax.parse;

import java.util.List;

import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;

import com.google.common.collect.Lists;

/**
 * 当行读取数据
 * @author JueYue
 * @param <T>
 * @date 2015年1月1日 下午7:59:39
 */
@SuppressWarnings("rawtypes")
public class SaxRowRead implements ISaxRowRead {

    private List list;

    public SaxRowRead() {
        list = Lists.newArrayList();
    }

    @Override
    public <T> List<T> getList() {
        return list;
    }

    @Override
    public void parse(int index, List<SaxReadCellEntity> datas) {

    }

}
