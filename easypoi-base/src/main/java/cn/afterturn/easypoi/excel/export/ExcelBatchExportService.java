package cn.afterturn.easypoi.excel.export;

import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.styler.IExcelExportStyler;
import cn.afterturn.easypoi.exception.excel.ExcelExportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelExportEnum;
import cn.afterturn.easypoi.handler.inter.IExcelExportServer;
import cn.afterturn.easypoi.util.PoiExcelGraphDataUtil;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.lang.reflect.Field;
import java.util.*;

import static cn.afterturn.easypoi.excel.ExcelExportUtil.USE_SXSSF_LIMIT;

/**
 * 提供批次插入服务
 *
 * @author JueYue
 * 2016年8月29日
 */
public class ExcelBatchExportService extends ExcelExportService {

    private Workbook                workbook;
    private Sheet                   sheet;
    private List<ExcelExportEntity> excelParams;
    private ExportParams            entity;
    private int                     titleHeight;
    private Drawing                 patriarch;
    private short                   rowHeight;
    private int                     index;

    public void init(ExportParams entity, Class<?> pojoClass) {
        List<ExcelExportEntity> excelParams = createExcelExportEntityList(entity, pojoClass);
        init(entity, excelParams);
    }

    public void init(ExportParams entity, List<ExcelExportEntity> excelParams) {
        LOGGER.debug("ExcelBatchExportServer only support SXSSFWorkbook");
        entity.setType(ExcelType.XSSF);
        workbook = new SXSSFWorkbook();
        this.entity = entity;
        this.excelParams = excelParams;
        super.type = entity.getType();
        createSheet(workbook, entity, excelParams);
        if (entity.getMaxNum() == 0) {
            entity.setMaxNum(USE_SXSSF_LIMIT);
        }
        insertDataToSheet(workbook, entity, excelParams, null, sheet);
    }

    public List<ExcelExportEntity> createExcelExportEntityList(ExportParams entity, Class<?> pojoClass) {
        try {
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            // 得到所有字段
            Field[]     fileds   = PoiPublicUtil.getClassFields(pojoClass);
            ExcelTarget etarget  = pojoClass.getAnnotation(ExcelTarget.class);
            String      targetId = etarget == null ? null : etarget.value();
            getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass,
                    null, null);
            sortAllParams(excelParams);

            return excelParams;
        } catch (Exception e) {
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }
    }

    public void createSheet(Workbook workbook, ExportParams entity, List<ExcelExportEntity> excelParams) {
        if (LOGGER.isDebugEnabled()) {
            LOGGER.debug("Excel export start ,List<ExcelExportEntity> is {}", excelParams);
            LOGGER.debug("Excel version is {}",
                    entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
        }
        if (workbook == null || entity == null || excelParams == null) {
            throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
        }
        try {
            try {
                sheet = workbook.createSheet(entity.getSheetName());
            } catch (Exception e) {
                // 重复遍历,出现了重名现象,创建非指定的名称Sheet
                sheet = workbook.createSheet();
            }
        } catch (Exception e) {
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }
    }

    public Workbook appendData(Collection<?> dataSet) {
        if (sheet.getLastRowNum() + dataSet.size() > entity.getMaxNum()) {
            sheet = workbook.createSheet();
            index = 0;
        }

        Iterator<?> its = dataSet.iterator();
        while (its.hasNext()) {
            Object t = its.next();
            try {
                index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight, 0)[0];
            } catch (Exception e) {
                LOGGER.error(e.getMessage(), e);
                throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
            }
        }
        return workbook;
    }

    @Override
    protected void insertDataToSheet(Workbook workbook, ExportParams entity,
                                     List<ExcelExportEntity> entityList, Collection<?> dataSet,
                                     Sheet sheet) {
        try {
            dataHandler = entity.getDataHandler();
            if (dataHandler != null && dataHandler.getNeedHandlerFields() != null) {
                needHandlerList = Arrays.asList(dataHandler.getNeedHandlerFields());
            }
            dictHandler = entity.getDictHandler();
            i18nHandler = entity.getI18nHandler();
            // 创建表格样式
            setExcelExportStyler((IExcelExportStyler) entity.getStyle()
                    .getConstructor(Workbook.class).newInstance(workbook));
            patriarch = PoiExcelGraphDataUtil.getDrawingPatriarch(sheet);
            List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
            if (entity.isAddIndex()) {
                excelParams.add(indexExcelEntity(entity));
            }
            excelParams.addAll(entityList);
            sortAllParams(excelParams);
            this.index = entity.isCreateHeadRows()
                    ? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
            titleHeight = index;
            setCellWith(excelParams, sheet);
            setColumnHidden(excelParams, sheet);
            rowHeight = getRowHeight(excelParams);
            setCurrentIndex(1);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
        }
    }

    public Workbook closeExportBigExcel() {
        if (entity.getFreezeCol() != 0) {
            sheet.createFreezePane(entity.getFreezeCol(), titleHeight, entity.getFreezeCol(), titleHeight);
        }
        mergeCells(sheet, excelParams, titleHeight);
        // 创建合计信息
        addStatisticsRow(getExcelExportStyler().getStyles(true, null), sheet);
        return workbook;
    }

    public Workbook exportBigExcel(IExcelExportServer server, Object queryParams) {
        int page = 1;
        List<Object> list = server
                .selectListForExcelExport(queryParams, page++);
        while (list != null && list.size() > 0) {
            appendData(list);
            list = server.selectListForExcelExport(queryParams, page++);
        }
        return closeExportBigExcel();
    }

}
