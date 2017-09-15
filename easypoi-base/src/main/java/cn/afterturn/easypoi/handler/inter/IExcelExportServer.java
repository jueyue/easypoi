package cn.afterturn.easypoi.handler.inter;

import java.util.List;

/**
 * 导出数据接口
 * @author JueYue
 * 2016年9月8日
 */
public interface IExcelExportServer {
    /**
     * 查询数据接口
     * @param obj    查询条件
     * @param page   当前页数
     * @return
     */
    public List<Object> selectListForExcelExport(Object obj, int page);

}
