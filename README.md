===========================
EasyPoi Excel和 Word简易工具类
===========================
 easypoi功能如同名字easy,主打的功能就是容易,让一个没见接触过poi的人员
就可以方便的写出Excel导出,Excel模板导出,Excel导入,Word模板导出,通过简单的注解和模板
语言(熟悉的表达式语法),完成以前复杂的写法

	作者博客：http://blog.afterturn.cn/
	作者邮箱： qrb.jueyue@gmail.com
	QQ群:  364192721
	
	开发者:魔幻之翼 xf.key@163.com



[开发文档请查看DOC下面的EasyPoi教程](http://git.oschina.net/jueyue/easypoi/blob/master/doc/EasyPoi%E6%95%99%E7%A8%8B.docx?dir=0&filepath=doc%2FEasyPoi%E6%95%99%E7%A8%8B.docx&oid=133ebde5641f128420023dbdc2533ab4b53c7792&sha=4c8d94c80544811e2e358518bad6c689a63c13e9)
!!!2.1.6 版本开始和之前的版本校验不兼用,使用JSR303的校验,删除了之前的注解,请注意
!!! 2.3.0 模板导出有问题,请使用2.3.0.1修复版本


版本介绍
history.md

基础示例
basedemo.md

[测试项目](http://git.oschina.net/jueyue/easypoi-test): http://git.oschina.net/jueyue/easypoi-test

**!!! 3.0.1 版本开始全新包名和GROUPID cn.afterturn**

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
			<version>2.4.0</version>
		</dependency>
		<dependency>
			<groupId>org.jeecg</groupId>
			<artifactId>easypoi-web</artifactId>
			<version>2.4.0</version>
		</dependency>
		<dependency>
			<groupId>org.jeecg</groupId>
			<artifactId>easypoi-annotation</artifactId>
			<version>2.4.0</version>
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

