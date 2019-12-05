
EasyPoi - Easy utility classes of Excel and Word
===========================

 Easypoi, as is clear from the name, it's easy for a developer who never even used poi to
export/import Excel, export Excel Template and Word Template, and export PDF. We encapsulate Apache poi in the upper layers. 
With simple annotations and templates Language (familiar expression syntax) to accomplish previously complex coding.

	Website：https://opensource.afterturn.cn/
	Email： qrb.jueyue@foxmail.com
	Developer:Jueyue qrb.jueyue@foxmail.com
	Excellent Team, undertake project development
Support Spring Boot    https://gitee.com/lemur/easypoi-spring-boot-starter

[Official website](https://opensource.afterturn.cn)

[中文介绍](https://gitee.com/lemur/easypoi/blob/master/README-cn.md)


**The Dev Guide**

**[https://opensource.afterturn.cn/doc/easypoi.html](http://opensource.afterturn.cn/doc/easypoi.html)**

Next steps
 - Internationalization, translate documents and code-commenting
 - Change all PDF export to Template export
 - Synchronise Word tmplate function with the Excel

[User feedback](https://github.com/jueyue/easypoi/issues/2)

## cn.afterturn:easypoi for enterprise

Available as part of the Tidelift Subscription

The maintainers of cn.afterturn:easypoi and thousands of other packages are working with Tidelift to deliver commercial support and maintenance for the open source dependencies you use to build your applications. Save time, reduce risk, and improve code health, while paying the maintainers of the exact dependencies you use. [Learn more.](https://tidelift.com/subscription/pkg/maven-cn-afterturn-easypoi?utm_source=maven-cn-afterturn-easypoi&utm_medium=referral&utm_campaign=readme)

## Security contact information

      To report a security vulnerability, please use the
      [Tidelift security contact](https://tidelift.com/security).
      Tidelift will coordinate the fix and disclosure.

Display in order of registration. If you are using Easypoi, please register on https://gitee.com/lemur/easypoi/issues/IFDX7 which is only as a reference for the open source, no other purposes.

   - [beyondsoft](http://www.beyondsoft.com)
   - [turingoal](http://www.turingoal.com)
   - [863soft](http://www.863soft.com/cn/)
   - [163](http://www.163.com)
   - [towngas](http://www.towngas.com.cn/)
   - [weifenghr](https://www.weifenghr.com/)
   - [ic-credit](http://sic-credit.cn/)
   - [getto1](https://www.getto1.com/)
   - [choicesoft](http://www.choicesoft.com.cn/)
   - [timeyaa](https://www.timeyaa.com/)
   
    

Version introduction

[history.md](https://gitee.com/lemur/easypoi/blob/master/history.md)

Basic demo

[basedemo.md](https://gitee.com/lemur/easypoi/blob/master/basedemo.md)

[Demo project](http://git.oschina.net/lemur/easypoi-test): http://git.oschina.net/lemur/easypoi-test

---------------------------
Advantages of EasyPoi
--------------------------
	1.Exquisite Design, easy to use
	2.Various Interfaces, easy to extend
	3.Coding less do more
	4.Support Spring MVC, easy for WEB export

---------------------------
Main Features
--------------------------

For Excel, self-adapt xls and xlsx format. For Word, only docx.

1.Excel Import
   - Annotation Import
   - Map Import
   - Big data Import, sax mode
   - Save file
   - File validation
   - Field validation

2.Excel Export
   - Annotation Export
   - Template Export
   - HTML Export

3.Excel convert to HTML

4.Word Export

5.PDF Export

---------------------------
which methods for which scenarios
---------------------------

    - Export
	    1. Normal excel export (simple format, moderate amount of data, within 50,000)
	        annotation way:  ExcelExportUtil.exportExcel(ExportParams entity, Class<?> pojoClass,Collection<?> dataSet) 
	    2. uncertain columns, but also with simple format and small amount of data
	        customize way: ExcelExportUtil.exportExcel(ExportParams entity, List<ExcelExportEntity> entityList,Collection<?> dataSet)
	    3. big data(greater than 50,000, less than one million) 
	        annotation way ExcelExportUtil.exportBigExcel(ExportParams entity, Class<?> pojoClass,IExcelExportServer server, Object queryParams)
	        customize way: ExcelExportUtil.exportBigExcel(ExportParams entity, List<ExcelExportEntity> excelParams,IExcelExportServer server, Object queryParams)
	    4. complex style, the amount of data not too large
	        Template: ExcelExportUtil.exportExcel(TemplateExportParams params, Map<String, Object> map)
	    5. Export multiple sheets with different styles at one time
	        Template ExcelExportUtil.exportExcel(Map<Integer, Map<String, Object>> map,TemplateExportParams params) 
	    6. One template but many copies to export
	        Template ExcelExportUtil.exportExcelClone(Map<Integer, List<Map<String, Object>>> map,TemplateExportParams params)
	    7. If Template can't satisfy your customization, try HTML
	        Build your own html, and then convert it to excel:  ExcelXorHtmlUtil.htmlToExcel(String html, ExcelType type)
	    8. Big data(Over millions), please use CSV
	        annotation way: CsvExportUtil.exportCsv(CsvExportParams params, Class<?> pojoClass, OutputStream outputStream)
	        customize way: CsvExportUtil.exportCsv(CsvExportParams params, List<ExcelExportEntity> entityList, OutputStream outputStream)
        9. Word Export
            Template: WordExportUtil.exportWord07(String url, Map<String, Object> map)
        10. PDF Export
            Template: TODO 
    - Import
        If you want to improve the performance, the concurrentTask of ImportParams will help with concurrent imports;Only support single line, minimum 1000. 
		For some special reading for single cell, you can use the  readSingleCell parameter
        1. no need verification, moderate amount of data, within 50,000
            annotation or map: ExcelImportUtil.importExcel(File file, Class<?> pojoClass, ImportParams params)
        2. the amount of data not too large
            annotation or map: ExcelImportUtil.importExcelMore(InputStream inputstream, Class<?> pojoClass, ImportParams params)
        3. For big data or with lots of importing operations; Less memory, only support single line
           SAX: ExcelImportUtil.importExcelBySax(InputStream inputstream, Class<?> pojoClass, ImportParams params, IReadHandler handler)
        4. Big data(Over millions), please use CSV
            small data: CsvImportUtil.importCsv(InputStream inputstream, Class<?> pojoClass,CsvImportParams params)
            big data: CsvImportUtil.importCsv(InputStream inputstream, Class<?> pojoClass,CsvImportParams params, IReadHandler readHandler)
        
	        
	
---------------------------
The difference between XLS and XLSX for Excel Export
---------------------------

	1. For export time, XLS is 2-3 times faster than xlsx.
	2. For export size, XLS is 2-3 or more times than xlsx.
	3. Export need to consider both network speed and local speed.
	
---------------------------
Packages Guide
---------------------------
	1.easypoi -- Parent package
	2.easypoi-annotation -- Basic annotation package, action on entity objects, 
                            it's convenient for Maven multi-project management after splitting
	3.easypoi-base -- Import and export package, realize Excel Import/Export, Word Export
	4.easypoi-web  -- Based on AbstractView, Coupled with spring MVC, greatly simplifies the export function 
	5.sax -- Optional, Export uses xercesImpl, Word Export uses poi-scratchpad

	If you don't use spring MVC, only easypoi-base is enough.
--------------------------
maven 
--------------------------
https://oss.sonatype.org/content/groups/public/
```xml
		 <dependency>
              <groupId>cn.afterturn</groupId>
              <artifactId>easypoi-spring-boot-starter</artifactId>
              <version>4.1.3</version>
         </dependency>
		 <dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-base</artifactId>
			<version>4.1.3</version>
		</dependency>
		<dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-web</artifactId>
			<version>4.1.3</version>
		</dependency>
		<dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-annotation</artifactId>
			<version>4.1.3</version>
		</dependency>
		
```
	

--------------------------
pom
--------------------------

```xml
			<!-- sax: optional, using for import -->
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
			
			<!-- Word: optional, using for export  -->
            <dependency>
                <groupId>org.apache.poi</groupId>
                <artifactId>ooxml-schemas</artifactId>
                <version>1.3</version>
                <optional>true</optional>
            </dependency>
			
			<!-- Verification: optional -->
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
			
			<!-- PDF: optional, using for export-->
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
Current test coverage 
----------------
|Package|Class|Method|Line|
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

