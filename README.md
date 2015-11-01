===========================
EasyPoi Excel和 Word简易工具类
===========================
 easypoi功能如同名字easy,主打的功能就是容易,让一个没见接触过poi的人员
就可以方便的写出Excel导出,Excel模板导出,Excel导入,Word模板导出,通过简单的注解和模板
语言(熟悉的表达式语法),完成以前复杂的写法

	作者博客：http://blog.afterturn.cn/
	作者邮箱： qrb.jueyue@gmail.com
	QQ群:  364192721

[测试项目](http://git.oschina.net/jueyue/easypoi-test): http://git.oschina.net/jueyue/easypoi-test

!!!2.1.6-SNAPSHOT 版本开始和之前的版本校验不兼用,使用hibernate的校验,删除了之前的注解,请注意
	
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

	1.ExcelExportUtil Excel导出(
	普通导出,模板导出)
	2.ExcelImportUtil Excel导入
	3.WordExportUtil Word导出(只支持docx ,doc版本poi存在图片的bug,暂不支持)
	
---------------------------
关于Excel导出XLS和XLSX区别
---------------------------

	1.导出时间XLS比XLSX快2-3倍
	2.导出大小XLS是XLSX的2-3倍或者更多
	3.导出需要综合网速和本地速度做考虑^~^
	
---------------------------
几个工程的说明
---------------------------
	1.easypoi 父包--作用大家都懂得
	2.easypoi-annotation 基础注解包,作用与实体对象上,拆分后方便maven多工程的依赖管理
	3.easypoi-base 导入导出的工具包,可以完成Excel导出,导入,Word的导出,Excel的导出功能
	4.easypoi-web  耦合了spring-mvc 基于AbstractView,极大的简化spring-mvc下的导出功能
	5.sax 导入使用xercesImpl这个包(这个包可能造成奇怪的问题哈),word导出使用poi-scratchpad,都作为可选包了
--------------------------
maven 
--------------------------
maven库应该都可以了
SNAPSHOT 版本
https://oss.sonatype.org/content/repositories/snapshots/
```xml
		 <dependency>
			<groupId>org.jeecg</groupId>
			<artifactId>easypoi-base</artifactId>
			<version>2.1.5</version>
		</dependency>
		<dependency>
			<groupId>org.jeecg</groupId>
			<artifactId>easypoi-web</artifactId>
			<version>2.1.5</version>
		</dependency>
		<dependency>
			<groupId>org.jeecg</groupId>
			<artifactId>easypoi-annotation</artifactId>
			<version>2.1.5</version>
		</dependency>
```
	
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

[EasyPoi 在spring mvc中简易开发方式](http://note.youdao.com/share/?id=8b04c4d88a0574a59aeaffd142c8e34b&type=note)

[EasyPoi-新版自定义导出样式类型](http://note.youdao.com/share/?id=7937a9fe15f1016b8f39bf813be894f8&type=note)

[EasyPoi-标签-模板导出的语法介绍](http://blog.csdn.net/qjueyue/article/details/45231801)

[EasyPoi-Excel预览(New)](http://blog.afterturn.cn/?p=34)

[EasyPoi-多行模板输出(New)](http://blog.afterturn.cn/?p=40)

[EasyPoi-校验集成](http://blog.afterturn.cn/?p=65)


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


--------------------------
pom说明
--------------------------
word和sax读取的时候才使用,就不是必须的了,请手动引用
```xml
			<!-- sax 读取时候用到的 -->
			<dependency>
				<groupId>xerces</groupId>
				<artifactId>xercesImpl</artifactId>
				<version>${xerces.version}</version>
				<optional>true</optional>
			</dependency>
			<dependency>
				<groupId>org.apache.poi</groupId>
				<artifactId>poi-scratchpad</artifactId>
				<version>${poi.version}</version>
				<optional>true</optional>
			</dependency>
```

--------------------------
版本修改
--------------------------

 - 2.1.5 bug修复,建议升级
 	- fix指定sheetNum导入的bug
 	- add导入时候setIndex(null) 可以防止识别为集合的问题
 	- fix导入excel列数量判断有bug，导致有些列数漏掉校验
 	- update修改了getValue的方式,减少导入cell value识别错误
 - 2.1.4
 	- 模板输出自动合并单元格功能
 	- 多行模板数据导出
 	- 导出链接功能
 	- 把反射加入了缓存
 	- word和Excel语法保持了统一
 	- Excel to Html 增加了图片显示和缓存功能
 	- 导入增加一个参数startSheetIndex 用以指定sheet位置

 - 2.1.3
 	- 屏蔽了科学计数法的问题
 	- 修复了多模版记得清空记录信息
 	- 加入了map setValue的接口,用以自定义setValue,主要是可以自定义key
 	- 把合并单元格的数据提起出来,让大家都可以复用
 	- 多个sheet导出问题

 - 2.1.2
 	- groupId 改成org.jeecg,为了以后可以直接提交到中央库了
 	- 把test 那个模块删除了
 	- 提供了Excel 预览的功能
 	- type 新增一个 4类型,用于Excel是数字,但是java 是String,防止科学计数法
 	- 修复一些bug
 - 2.1.1-release -小更新/防止下一个预览功能太久不能发布更新
 	- $fe标签
 	- 几个可能影响使用的bug
 	- 导入Map的支持
 - 2.0.9-release--模板功能更新
 	- 修复了中文路径,float类型导入等bug
 	- 增加了根据CellStyler判断cell类型的功能
 	- 优化了一下代码,sort 的compare 改为类实现接口,getValue统一由public处理
 	- height和width 都改成double 可能更加准确的调整
 	- 升级到common-lang3
 	- 可以循环解析多个模板
 	- !!模板增加了多个标签功能
 - 2.0.8-release--小版本更新
 	- 分开了基础注解和base包,编译maven多模块集成
 	- 加强了Excel导入的校验功能,可以追加错误信息,过滤不合格数据
 	- 修复了spring mvc下的07版本不支持的问题
 	- 添加了@  Controller 注解,扫描org.jeecgframework.poi.excel.view路径就可以了,不用写bean了
 	- 加强了styler的自定义功能,参数改为entity,自由控制
 	- test包下面增加了几个demo
 	- 导出添加了一个冰冻列属性,可以简单执行
 	- PS: 经过这段时间项目中的测试,模板导出Excel复杂报表可以省5倍以上的时间,特别是样式复杂的完全可以在Excel中完成,不是在代码中完成了
 - 2.0.7-release--推荐更新
 	- 增加了合计的参数,便于统计合计信息
 	- 修改了样式设置,使用默认设置,提供其他设置,和样式接口
 	- 模板导出增加了插入导出的功能
 	- 模板导出重用了Excel导出的代码功能更加强劲
 	- 新增了Sax导入方式,大数据导入提供接口操作
 	- Word修改了页眉页脚不替换的bug
 	- 修改了捕获异常日志的bug
 	- 修复了模板导出,Spring View没识别版本的bug
 	- 修复了其他一些小bug
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
	/**
     * Email校验
     */
    @Excel(name = "Email", width = 25)
    @ExcelVerify(isEmail = true, notNull = true)
    private String email;
    /**
     * 手机号校验
     */
    @Excel(name = "Mobile", width = 20)
    @ExcelVerify(isMobile = true, notNull = true)
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
