
EasyPoi Excel和 Word简易工具类
===========================

 easypoi功能如同名字easy,主打的功能就是容易,让一个没见接触过poi的人员
就可以方便的写出Excel导出,Excel模板导出,Excel导入,Word模板导出,通过简单的注解和模板
语言(熟悉的表达式语法),完成以前复杂的写法

	官网： http://opensource.afterturn.cn/
	邮箱： xue315998@qq.com
	QQ群: 1群 364192721(满) 2 群116844390
	
	开发者:魔幻之翼 xf.key@163.com
	优秀团队,承接私活,加微信:13165188222 ,添加备注悟耘
Spring Boot 支持    https://gitee.com/lemur/easypoi-spring-boot-starter

[官网](https://opensource.afterturn.cn)


**开发指南**

**[https://opensource.afterturn.cn/doc/easypoi.html](http://opensource.afterturn.cn/doc/easypoi.html)**

下一步计划
 - 国际化,翻译文档以及注释
 - pdf全部改为模板导出方式
 - word模板和excel功能同步

[用户征集](https://gitee.com/lemur/easypoi/issues/IFDX7)


按登记顺序展示,如果你正在使用EasyPoi欢迎在 https://gitee.com/lemur/easypoi/issues/IFDX7 登记。仅作为开源参考，不作其它用途

   - [一众科技](http://studio.afterturn.cn/)
   - [博丰(深圳)跨境电商服务有限公司 ](http://www.bfeservice.com/)
   - [陕西昊景无线网络技术有限公司](http://www.hjwxbd.com)
   - [湖北省楚天云有限公司](http://www.chutianyun.gov.cn)
   - [博彦科技](http://www.beyondsoft.com)
   - [图灵谷北京科技有限公司](http://www.turingoal.com)
   - [多贝拉](https://www.duobeila.cn/)
   - [河南八六三软件股份有限公司](http://www.863soft.com/cn/)
   - [网易](http://www.netease.com)
   - [港华燃气](http://www.towngas.com.cn/)
   - [北京易才博普奥管理顾问有限公司](https://www.weifenghr.com/)
   - [中投国信(北京)科技发展有限公司](http://sic-credit.cn/)
   - [武汉格事化信息技术](https://www.getto1.com/)
   - [北京辰森世纪科技股份有限公司](http://www.choicesoft.com.cn/)
   
    

版本介绍

[history.md](https://gitee.com/lemur/easypoi/blob/master/history.md)

基础示例

[basedemo.md](https://gitee.com/lemur/easypoi/blob/master/basedemo.md)

[测试项目](http://git.oschina.net/lemur/easypoi-test): http://git.oschina.net/lemur/easypoi-test

---------------------------
EasyPoi的主要特点
--------------------------
	1.设计精巧,使用简单
	2.接口丰富,扩展简单
	3.默认值多,write less do more
	4.AbstractView 支持,web导出可以简单明了

---------------------------
什么场景该用哪个方法
---------------------------

    - 导出
	    1.正规excel导出 (格式简单,数据量可以,5W以内吧)
	        注解方式:  ExcelExportUtil.exportExcel(ExportParams entity, Class<?> pojoClass,Collection<?> dataSet) 
	    2.不定多少列,但是格式依然简单数据库不大
	        自定义方式: ExcelExportUtil.exportExcel(ExportParams entity, List<ExcelExportEntity> entityList,Collection<?> dataSet)
	    3.数据量大超过5W,还在100W以内
	        注解方式 ExcelExportUtil.exportBigExcel(ExportParams entity, Class<?> pojoClass,IExcelExportServer server, Object queryParams)
	        自定义方式: ExcelExportUtil.exportBigExcel(ExportParams entity, List<ExcelExportEntity> excelParams,IExcelExportServer server, Object queryParams)
	    4.样式复杂,数据量尽量别大
	        模板导出 ExcelExportUtil.exportExcel(TemplateExportParams params, Map<String, Object> map)
	    5.一次导出多个风格不一致的sheet
	        模板导出 ExcelExportUtil.exportExcel(Map<Integer, Map<String, Object>> map,TemplateExportParams params) 
	    6.一个模板但是要导出非常多份
	        模板导出 ExcelExportUtil.exportExcelClone(Map<Integer, List<Map<String, Object>>> map,TemplateExportParams params)
	    7.模板无法满足你的自定义,试试html
	        自己构造html,然后我给你转成excel  ExcelXorHtmlUtil.htmlToExcel(String html, ExcelType type)
	    8.数据量过百万级了.放弃excel吧,csv导出
	        注解方式: CsvExportUtil.exportCsv(CsvExportParams params, Class<?> pojoClass, OutputStream outputStream)
	        自定义方式: CsvExportUtil.exportCsv(CsvExportParams params, List<ExcelExportEntity> entityList, OutputStream outputStream)
        9.word导出
            模板导出: WordExportUtil.exportWord07(String url, Map<String, Object> map)
        10.PDF导出
            模板导出: TODO 
    - 导入 
        如果想提高性能 ImportParams 的concurrentTask 可以帮助并发导入,仅单行,最小1000
        excel有单个的那种特殊读取,readSingleCell 参数可以支持
        1. 不需要检验,数据量不大(5W以内)
            注解或者MAP: ExcelImportUtil.importExcel(File file, Class<?> pojoClass, ImportParams params)
        2. 需要导入,数据量不大
            注解或者MAP: ExcelImportUtil.importExcelMore(InputStream inputstream, Class<?> pojoClass, ImportParams params)
        3. 数据量大了,或者你有特别多的导入操作,内存比较少,仅支持单行
           SAX方式  ExcelImportUtil.importExcelBySax(InputStream inputstream, Class<?> pojoClass, ImportParams params, IReadHandler handler)
        4. 数据量超过EXCEL限制,CSV读取
            小数据量: CsvImportUtil.importCsv(InputStream inputstream, Class<?> pojoClass,CsvImportParams params)
            大数据量: CsvImportUtil.importCsv(InputStream inputstream, Class<?> pojoClass,CsvImportParams params, IReadHandler readHandler)
        
	        
	
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
SNAPSHOT 版本-很少发布
https://oss.sonatype.org/content/groups/public/
```xml
		 <dependency>
              <groupId>cn.afterturn</groupId>
              <artifactId>easypoi-spring-boot-starter</artifactId>
              <version>4.1.2</version>
         </dependency>
		 <dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-base</artifactId>
			<version>4.1.2</version>
		</dependency>
		<dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-web</artifactId>
			<version>4.1.2</version>
		</dependency>
		<dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-annotation</artifactId>
			<version>4.1.2</version>
		</dependency>
		
```
	

--------------------------
pom说明
--------------------------
word和sax读取的时候才使用,就不是必须的了,请手动引用,JSR303的校验也是可选的,PDF的jar也是可选的
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
			
			<!-- Word 需要使用 -->
            <dependency>
                <groupId>org.apache.poi</groupId>
                <artifactId>ooxml-schemas</artifactId>
                <version>1.3</version>
                <optional>true</optional>
            </dependency>
			
			<!-- 校验,下面两个实现 -->
			<dependency>
				<groupId>org.hibernate</groupId>
				<artifactId>hibernate-validator</artifactId>
				<version>5.1.3.Final</version>
				<optional>true</optional>
			</dependency>
			
			<dependency>
				<groupId>org.apache.bval</groupId>
				<artifactId>org.apache.bval.bundle</artifactId>
				<version>1.1.0</version>
			</dependency>
			
			<!-- PDF -->
			<dependency>
				<groupId>com.itextpdf</groupId>
				<artifactId>itextpdf</artifactId>
				<version>5.5.6</version>
				<optional>true</optional>
			</dependency>

			<dependency>
				<groupId>com.itextpdf</groupId>
				<artifactId>itext-asian</artifactId>
				<version>5.2.0</version>
				<optional>true</optional>
			</dependency>
```
-----------------
Test 测试覆盖率
----------------
|包|类|方法|行|
|----|----|----|----|
cn|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi|0% (0/1)|0% (0/1)|0% (0/3)
cn.afterturn.easypoi.cache|100% (6/6)|83% (10/12)|60% (40/66)
cn.afterturn.easypoi.cache.manager|100% (3/3)|62% (5/8)|54% (29/53)
cn.afterturn.easypoi.configuration|0% (0/2)|0% (0/1)|0% (0/5)
cn.afterturn.easypoi.entity|100% (1/1)|83% (15/18)|66% (26/39)
cn.afterturn.easypoi.entity.vo|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.excel|100% (3/3)|73% (19/26)|59% (44/74)
cn.afterturn.easypoi.excel.annotation|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.excel.entity|83% (5/6)|66% (85/127)|63% (185/290)
cn.afterturn.easypoi.excel.entity.enmus|100% (3/3)|88% (8/9)|92% (13/14)
cn.afterturn.easypoi.excel.entity.params|100% (6/6)|84% (94/111)|75% (160/211)
cn.afterturn.easypoi.excel.entity.result|100% (2/2)|76% (16/21)|65% (26/40)
cn.afterturn.easypoi.excel.entity.sax|100% (1/1)|33% (2/6)|45% (5/11)
cn.afterturn.easypoi.excel.entity.vo|100% (1/1)|100% (1/1)|100% (3/3)
cn.afterturn.easypoi.excel.export|100% (2/2)|94% (17/18)|84% (202/238)
cn.afterturn.easypoi.excel.export.base|100% (2/2)|97% (33/34)|88% (353/399)
cn.afterturn.easypoi.excel.export.styler|100% (4/4)|86% (19/22)|97% (122/125)
cn.afterturn.easypoi.excel.export.template|100% (3/3)|95% (43/45)|92% (422/457)
cn.afterturn.easypoi.excel.graph|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.excel.graph.builder|0% (0/1)|0% (0/8)|0% (0/92)
cn.afterturn.easypoi.excel.graph.constant|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.excel.graph.entity|0% (0/3)|0% (0/26)|0% (0/49)
cn.afterturn.easypoi.excel.html|100% (2/2)|92% (24/26)|84% (242/287)
cn.afterturn.easypoi.excel.html.css|100% (2/2)|100% (10/10)|94% (166/175)
cn.afterturn.easypoi.excel.html.css.impl|100% (6/6)|60% (9/15)|74% (111/149)
cn.afterturn.easypoi.excel.html.entity|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.excel.html.entity.style|100% (3/3)|100% (53/53)|100% (87/87)
cn.afterturn.easypoi.excel.html.helper|80% (4/5)|77% (24/31)|67% (154/228)
cn.afterturn.easypoi.excel.imports|100% (2/2)|92% (26/28)|79% (361/454)
cn.afterturn.easypoi.excel.imports.base|100% (1/1)|100% (9/9)|89% (102/114)
cn.afterturn.easypoi.excel.imports.sax|100% (2/2)|100% (8/8)|70% (56/80)
cn.afterturn.easypoi.excel.imports.sax.parse|100% (1/1)|77% (7/9)|58% (44/75)
cn.afterturn.easypoi.exception|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.exception.excel|50% (1/2)|20% (3/15)|18% (6/32)
cn.afterturn.easypoi.exception.excel.enums|50% (1/2)|44% (4/9)|40% (9/22)
cn.afterturn.easypoi.exception.word|0% (0/1)|0% (0/3)|0% (0/6)
cn.afterturn.easypoi.exception.word.enmus|0% (0/1)|0% (0/4)|0% (0/10)
cn.afterturn.easypoi.handler|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.handler.impl|100% (1/1)|33% (2/6)|44% (4/9)
cn.afterturn.easypoi.handler.inter|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.hanlder|0% (0/1)|0% (0/1)|0% (0/17)
cn.afterturn.easypoi.pdf|100% (1/1)|50% (1/2)|33% (1/3)
cn.afterturn.easypoi.pdf.entity|100% (1/1)|52% (10/19)|47% (17/36)
cn.afterturn.easypoi.pdf.export|100% (1/1)|100% (14/14)|85% (131/153)
cn.afterturn.easypoi.pdf.styler|100% (1/1)|100% (4/4)|56% (9/16)
cn.afterturn.easypoi.test|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.test.en|100% (3/3)|93% (15/16)|93% (31/33)
cn.afterturn.easypoi.test.entity|91% (11/12)|50% (62/122)|57% (117/202)
cn.afterturn.easypoi.test.entity.check|100% (2/2)|50% (8/16)|69% (18/26)
cn.afterturn.easypoi.test.entity.groupname|75% (3/4)|46% (25/54)|52% (48/92)
cn.afterturn.easypoi.test.entity.img|100% (1/1)|100% (8/8)|100% (16/16)
cn.afterturn.easypoi.test.entity.samename|100% (1/1)|100% (6/6)|100% (10/10)
cn.afterturn.easypoi.test.entity.statistics|100% (1/1)|100% (12/12)|100% (19/19)
cn.afterturn.easypoi.test.entity.temp|80% (4/5)|94% (32/34)|38% (52/134)
cn.afterturn.easypoi.test.web|100% (0/0)|100% (0/0)|100% (0/0)
cn.afterturn.easypoi.test.web.cfg|0% (0/1)|0% (0/1)|0% (0/4)
cn.afterturn.easypoi.tohtml|0% (0/1)|0% (0/4)|0% (0/13)
cn.afterturn.easypoi.util|100% (11/11)|85% (84/98)|73% (668/908)
cn.afterturn.easypoi.view|0% (0/13)|0% (0/27)|0% (0/376)
cn.afterturn.easypoi.word|100% (1/1)|33% (1/3)|20% (1/5)
cn.afterturn.easypoi.word.entity|100% (2/2)|44% (4/9)|45% (27/60)
cn.afterturn.easypoi.word.entity.params|50% (1/2)|22% (5/22)|14% (8/56)
cn.afterturn.easypoi.word.parse|100% (1/1)|91% (11/12)|94% (97/103)
cn.afterturn.easypoi.word.parse.excel|100% (2/2)|90% (10/11)|70% (95/135)

