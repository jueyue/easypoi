
EasyPoi Excel & Word Utils
----------------------------
 Easypoi functions like the name easy, the flagship feature is easy to let a person who has not seen the POI contact
 You can easily write Excel export, Excel template export, Excel import, Word template export, and through simple annotations and templates
 Language (familiar expressions, grammar), complete the previous complex writing

	site：http://www.afterturn.cn/
	mail： qrb.jueyue@gmail.com
	QQ群:  364192721
	
	dev: xf.key@163.com

**开发指南-中文[http://www.afterturn.cn/doc/easypoi.html](http://www.afterturn.cn/doc/easypoi.html)**


[base demo](https://github.com/lemur-open/easypoi/blob/master/basedemo.md)


[testProject](http://git.oschina.net/jueyue/easypoi-test): http://git.oschina.net/jueyue/easypoi-test


---------------------------
The main features of EasyPoi
--------------------------
1. exquisite design, easy to use
2. interface rich, easy to expand
3. default values are many, write, less, do, more
4. AbstractView support, web export can be simple and clear
---------------------------
Several entry tool classes for EasyPoi
---------------------------
1. ExcelExportUtil Excel export (General export, template export)
2. ExcelImportUtil Excel import
3. WordExportUtil Word export (support only docx, Doc version POI exist picture of bug, do not support)
---------------------------
About Excel export XLS and XLSX distinction
---------------------------
1. export time XLS is 2-3 times faster than XLSX
2. export size XLS is 2-3 times or more than XLSX
3. export needs comprehensive speed and speed do consider local ^ ~ ^
---------------------------
Description of several works
---------------------------
1. easypoi father package - everyone knows
2. easypoi-annotation foundation annotation package, role and entity object, after splitting convenient Maven multi project dependency management
3. easypoi-base import and export toolkit, you can complete the Excel export, import, export Word, Excel export function
4. easypoi-web coupled with spring-mvc, based on AbstractView, greatly simplifying the export function under spring-mvc
The 5.sax import uses the xercesImpl package (this package can cause a strange problem), and the word export uses poi-scratchpad as an optional package
--------------------------
maven 
--------------------------

```xml
		 <dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-base</artifactId>
			<version>3.0.1</version>
		</dependency>
		<dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-web</artifactId>
			<version>3.0.1</version>
		</dependency>
		<dependency>
			<groupId>cn.afterturn</groupId>
			<artifactId>easypoi-annotation</artifactId>
			<version>3.0.1</version>
		</dependency>
```
	

--------------------------
pom desc
--------------------------
Word and sax are used only when they are read. They are not necessary. Please refer to them manually. JSR303's check is optional, and PDF's jar is optional
```xml
			<!-- sax  -->
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
			
			<!-- Word  -->
            <dependency>
                <groupId>org.apache.poi</groupId>
                <artifactId>ooxml-schemas</artifactId>
                <version>1.3</version>
                <optional>true</optional>
            </dependency>
			
			<!-- validator -->
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

