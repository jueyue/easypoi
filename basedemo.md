
---------------------------
EasyPoi 文档
---------------------------
这几篇是旧的教程,不过和现在大同小异

[Excel 介绍篇](http://note.youdao.com/share/?id=ea5166076313d437abbfb28a18d85486&type=note)
[Excel 工具类](http://note.youdao.com/share/?id=dcfef5355d1d75b718be8fd563ad5216&type=note)
[Excel 注解介绍.第一篇](http://note.youdao.com/share/?id=9898d1777ca97b81d82a6e151b880dbb&type=note)
[Excel 注解介绍.第二篇](http://note.youdao.com/share/?id=9898d1777ca97b81d82a6e151b880dbb&type=note)
[Excel 实体类](http://note.youdao.com/share/?id=dab97eababf8e91356f87be204b581e8&type=note)
[Word模板导出教程](http://note.youdao.com/share/?id=26794c8eb4a285828663178c0ae854a2&type=note)

后面都是新的了

[EasyPoi 在spring mvc中简易开发方式](http://note.youdao.com/share/?id=8b04c4d88a0574a59aeaffd142c8e34b&type=note)

[EasyPoi-新版自定义导出样式类型](http://note.youdao.com/share/?id=7937a9fe15f1016b8f39bf813be894f8&type=note)

[EasyPoi-标签-模板导出的语法介绍](http://blog.csdn.net/qjueyue/article/details/45231801)



--------------------------
EasyPoi 模板 表达式支持
--------------------------
- 空格分割
- 三目运算  {{test ? obj:obj2}}
- n: 表示 这个cell是数值类型 {{n:}}
- le: 代表长度{{le:()}} 在if/else 运用{{le:() > 8 ? obj1 :  obj2}}
- fd: 格式化时间 {{fd:(obj;yyyy-MM-dd)}}
- fn: 格式化数字 {{fn:(obj;###.00)}}
- fe: 遍历数据,创建row
- !fe: 遍历数据不创建row 
- $fe: 下移插入,把当前行,下面的行全部下移.size()行,然后插入
- !if: 删除当前列 {{!if:(test)}}
- 单引号表示常量值 ''  比如'1' 那么输出的就是 1
- &NULL& 控制
- ]] 换行符



---------------------------
EasyPoi导出实例
---------------------------
1.注解,导入导出都是基于注解的,实体上做上注解,标示导出对象,同时可以做一些操作
```Java
	@ExcelTarget("courseEntity")
	public class CourseEntity implements java.io.Serializable {
	/** 主键 */
	private String id;
	/** 课程名称 */
	@Excel(name = "课程名称", orderNum = "1", needMerge = true)
	private String name;
	/** 老师主键 */
	@ExcelEntity(id = "yuwen")
	@ExcelVerify()
	private TeacherEntity teacher;
	/** 老师主键 */
	@ExcelEntity(id = "shuxue")
	private TeacherEntity shuxueteacher;

	@ExcelCollection(name = "选课学生", orderNum = "4")
	private List<StudentEntity> students;
```
2.基础导出
	传入导出参数,导出对象,以及对象列表即可完成导出
```Java
	HSSFWorkbook workbook = ExcelExportUtil.exportExcel(new ExportParams(
				"2412312", "测试", "测试"), CourseEntity.class, list);
```
3.基础导出,带有索引
	在到处参数设置一个值,就可以在导出列增加索引
```Java
	ExportParams params = new ExportParams("2412312", "测试", "测试");
	params.setAddIndex(true);
	HSSFWorkbook workbook = ExcelExportUtil.exportExcel(params,
			TeacherEntity.class, telist);
```			
4.导出Map
	创建类似注解的集合,即可完成Map的导出,略有麻烦
```Java
	List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
	entity.add(new ExcelExportEntity("姓名", "name"));
	entity.add(new ExcelExportEntity("性别", "sex"));

	List<Map<String, String>> list = new ArrayList<Map<String, String>>();
	Map<String, String> map;
	for (int i = 0; i < 10; i++) {
		map = new HashMap<String, String>();
		map.put("name", "1" + i);
		map.put("sex", "2" + i);
		list.add(map);
	}

	HSSFWorkbook workbook = ExcelExportUtil.exportExcel(new ExportParams(
			"测试", "测试"), entity, list);	
```			
5.模板导出
	根据模板配置,完成对应导出
```Java
	TemplateExportParams params = new TemplateExportParams();
	params.setHeadingRows(2);
	params.setHeadingStartRow(2);
	Map<String,Object> map = new HashMap<String, Object>();
    map.put("year", "2013");
    map.put("sunCourses", list.size());
    Map<String,Object> obj = new HashMap<String, Object>();
    map.put("obj", obj);
    obj.put("name", list.size());
	params.setTemplateUrl("org/jeecgframework/poi/excel/doc/exportTemp.xls");
	Workbook book = ExcelExportUtil.exportExcel(params, CourseEntity.class, list,
			map);
```			
6.导入
	设置导入参数,传入文件或者流,即可获得相应的list
```Java
	ImportParams params = new ImportParams();
	params.setTitleRows(2);
	params.setHeadRows(2);
	//params.setSheetNum(9);
	params.setNeedSave(true);
	long start = new Date().getTime();
	List<CourseEntity> list = ExcelImportUtil.importExcel(new File(
			"d:/tt.xls"), CourseEntity.class, params);
```	

7.和spring mvc的无缝融合
	简单几句话,Excel导出搞定
```Java
	@RequestMapping(params = "exportXls")
	public String exportXls(CourseEntity course,HttpServletRequest request,HttpServletResponse response
			, DataGrid dataGrid,ModelMap map) {

        CriteriaQuery cq = new CriteriaQuery(CourseEntity.class, dataGrid);
        org.jeecgframework.core.extend.hqlsearch.HqlGenerateUtil.installHql(cq, course, request.getParameterMap());
        List<CourseEntity> courses = this.courseService.getListByCriteriaQuery(cq,false);

        map.put(NormalExcelConstants.FILE_NAME,"用户信息");
        map.put(NormalExcelConstants.CLASS,CourseEntity.class);
        map.put(NormalExcelConstants.PARAMS,new ExportParams("课程列表", "导出人:Jeecg",
                "导出信息"));
        map.put(NormalExcelConstants.DATA_LIST,courses);
        return NormalExcelConstants.JEECG_EXCEL_VIEW;

	}
```

8.Excel导入校验,过滤不符合规则的数据,追加错误信息到Excel,提供常用的校验规则,已经通用的校验接口
```Java
    @Excel(name = "Email", width = 25)
    @Max(value = 15,message = "max 最大值不能超过15")
    private int email;
    /**
     * 手机号
     */
    @Excel(name = "Mobile", width = 20)
    @NotNull
    private String mobile;
    
    ExcelImportResult<ExcelVerifyEntity> result = ExcelImportUtil.importExcelVerify(new File(
            "d:/tt.xls"), ExcelVerifyEntity.class, params);
    for (int i = 0; i < result.getList().size(); i++) {
        System.out.println(ReflectionToStringBuilder.toString(result.getList().get(i)));
    }
```

9.导入Map
	设置导入参数,传入文件或者流,即可获得相应的list,自定义Key,需要实现IExcelDataHandler接口
```Java
	ImportParams params = new ImportParams();
	List<Map<String,Object>> list = ExcelImportUtil.importExcel(new File(
			"d:/tt.xls"), Map.class, params);
```	
10.大数据量Excel导出
	exportBigExcel 的方法 ,最后可以关闭closeExportBigExcel 也可以不关闭
	
11.如果View不起作用,已经发现被其他View处理掉的情况,使用下面这个,进行了统一封装,相同的效果
```Java
	PoiBaseView
	public static void render(Map<String, Object> model, HttpServletRequest request,
                              HttpServletResponse response, String viewName) {
        PoiBaseView view = null;
        if (BigExcelConstants.BIG_EXCEL_VIEW.equals(viewName)) {
            view = new BigExcelExportView();
        } else if (MapExcelConstants.JEECG_MAP_EXCEL_VIEW.equals(viewName)) {
            view = new JeecgMapExcelView();
        } else if (NormalExcelConstants.JEECG_EXCEL_VIEW.equals(viewName)) {
            view = new JeecgSingleExcelView();
        } else if (TemplateExcelConstants.JEECG_TEMPLATE_EXCEL_VIEW.equals(viewName)) {
            view = new JeecgTemplateExcelView();
        } else if (MapExcelGraphConstants.MAP_GRAPH_EXCEL_VIEW.equals(viewName)) {
            view = new MapGraphExcelView();
        }
        try {
            view.renderMergedOutputModel(model, request, response);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
    }
	// Demo
	@RequestMapping(params = "exportXls")
	public void exportXls(CourseEntity course,HttpServletRequest request,HttpServletResponse response
			, DataGrid dataGrid,ModelMap map) {
        CriteriaQuery cq = new CriteriaQuery(CourseEntity.class, dataGrid);
        org.jeecgframework.core.extend.hqlsearch.HqlGenerateUtil.installHql(cq, course, request.getParameterMap());
        List<CourseEntity> courses = this.courseService.getListByCriteriaQuery(cq,false);
        map.put(NormalExcelConstants.FILE_NAME,"用户信息");
        map.put(NormalExcelConstants.CLASS,CourseEntity.class);
        map.put(NormalExcelConstants.PARAMS,new ExportParams("课程列表", "导出人:Jeecg",
                "导出信息"));
        map.put(NormalExcelConstants.DATA_LIST,courses);
        PoiBaseView.render(map,request,response,NormalExcelConstants.JEECG_EXCEL_VIEW);
	}
```	
