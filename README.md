===========================
EasyPoi Excel和 Word简易工具类
===========================
 easypoi功能如同名字easy,主打的功能就是易容,让一个没见接触过poi的人员
就可以方便的写出Excel导出,Excel模板导出,Excel导入,Word模板导出,通过简单的注解和模板
语言(熟悉的表达式语法),完成以前复杂的写法

    作者博客：http://blog.csdn.net/qjueyue
	作者邮箱： qrb.jueyue@gmail.com

---------------------------
EasyPoi的主要特点
--------------------------
	1.设计精巧,使用简单
	2.接口丰富,扩展简单
	3.默认值多,write less do more
	4.AbstractView 支持,web导出可以简单明了

---------------------------
EasyPoi的几个入口工具类
---------------------------

	1.ExcelExportUtil Excel导出(普通导出,模板导出)
	2.ExcelImportUtil Excel导入
	3.WordExportUtil Word导出(只支持docx ,doc版本poi存在图片的bug,暂不支持)
	
---------------------------
关于Excel导出XLS和XLSX区别
---------------------------

	1.导出时间XLS比XLSX快2-3倍
	2.导出大小XLS是XLSX的2-3倍或者更多
	3.导出需要综合网速和本地速度做考虑^~^
	
--------------------------
版本修改
--------------------------
 - 2.0.6-release
 	- 修复导入的自定义格式异常
 	- 修复导入的BigDecimal不识别
 	- 导出大数据替换为SXSSFWorkbook,提高07版本Excel导出效率
 	- 根据sonar的提示,修复编码格式问题
 	- 做了更加丰富的测试
 	- word 导出的优化
 		
 - 2.0.6-SNAPSHOT
	 - 增加map的导出
	 - 增加index 列


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

[EasyPoi-在财务报表中的应用](http://git.oschina.net/jueyue/easypoi/wikis/EasyPoi-%E5%9C%A8%E8%B4%A2%E5%8A%A1%E6%8A%A5%E8%A1%A8%E4%B8%AD%E7%9A%84%E5%BA%94%E7%94%A8)

[EasyPoi-如何筛选导出属性](http://git.oschina.net/jueyue/easypoi/wikis/EasyPoi-%E5%A6%82%E4%BD%95%E7%AD%9B%E9%80%89%E5%AF%BC%E5%87%BA%E5%B1%9E%E6%80%A7)


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