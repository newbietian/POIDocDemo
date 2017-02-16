最近在项目中要生成Word的doc和docx文件，一番百度google之后，发现通过java语言实现的主流是Apache的POI组件。除了POI，这里还有[另一种实现](http://stackoverflow.com/questions/203174/whats-a-good-java-api-for-creating-word-documents)，不过我没有去研究，有兴趣的同学可以研究研究。

关于**POI**可以访问[Apache POI的官网](https://poi.apache.org/index.html)获取详细的信息。

进入主题！ 

由于项目中只是用到了doc和docx的组件，下面也只是介绍这两个组件的使用

###**一、在Android Studio中如何用POI组件**
从POI官网上看，貌似暂并不支持IntelliJ IDE，如下图，所以这里我们采用直接下载jar包并导入项目的方式。

![官网how to build](http://upload-images.jianshu.io/upload_images/3971774-d224bdeecc996120?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
 
通过[官网](https://poi.apache.org/index.html) ->Overview->Components，可以看到 d和docx文件分别对应着组件**HWPF**和**XWPF**，而HWPF和XWPF则对应着poi-scratchpad和poi-ooxml

| 文件类型 | 组件名| MavenId |
|:---|:----|:----|
|doc|HWPF|poi-scratchpad|
|docx|XWPF|poi-ooxml|

![Components Map](http://upload-images.jianshu.io/upload_images/3971774-3291b25bc144d05e?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)


----------


####**下载**
进入Apache[下载页面](https://poi.apache.org/download.html)，选择最新版下载，如下。选择The latest beta release is Apache POI 3.16-beta2会跳转到poi-bin-3.16-beta2-20170202.tar.gz，然后点击poi-bin-3.16-beta2-20170202.tar.gz，选择镜像后即可成功下载。

**注**
linux系统选择.tar.gz
windows系统选择.zip

![POI下载页面](http://upload-images.jianshu.io/upload_images/3971774-75b9acec7b6a5fb6?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)


----------


####**解压**
将下载后的压缩包解压，会得到以下文件。

|文件（夹）名 | 作用 |
| :- | :- |
|docs|文档（包括API文档和如何使用及版本信息）|
|lib|doc功能实现依赖的包|
|ooxml-lib|docx功能实现依赖的包|
|LICENSE||
|NOTICE||
|poi-3.16-beta2.jar|The prerequisite poi-scratchpad-3.16-beta2.jar|
|poi-examples-3.16-beta2.jar|不明确|
|poi-excelant-3.16-beta2.jar|excel功能实现|
|poi-ooxml-3.16-beta2.jar|docx功能实现|
|poi-ooxml-schemas-3.16-beta2.jar|The prerequisite of poi-ooxml-3.16-beta2.jar|
|poi-scratchpad-3.16-beta2.jar|doc功能实现|
  
  
![解压后的poi包](http://upload-images.jianshu.io/upload_images/3971774-3d6813696f0554fd?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)


----------

####**导入**
不熟悉怎么导入的同学可以看看[Android Studio导入jar包教程](http://blog.csdn.net/ygd1994/article/details/51346984)
1、doc
对于doc文件，需要将**lib文件夹下jar包，poi-3.16-beta2.jar，poi-scratchpad-3.16-beta2.jar**放入android项目libs目录下（lib文件夹下的junit-4.12.jar和log4j-1.2.17.jar不放我的项目也没出现异常，能少点是点）。

![项目中的libs](http://upload-images.jianshu.io/upload_images/3971774-f218f127cb5b3228?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

2、docx
对于docx，需要导入**lib文件夹下jar包，poi-3.16-beta2.jar，poi-ooxml-3.16-beta2.jar，poi-ooxml-schemas-3.16-beta2.jar和ooxml-lib下的包**，由于一直我这一直出现**Warning:Ingoring InnerClasses attribute for an anonymous inner class**的错误，同时由于doc基本满足我的需求以及导入这么多jar导致apk体积增大，就没有去实现。
有兴趣的同学可以研究研究。

    
           
           
----------------------------------------
### **二、实现doc文件的读写**

Apache POI中的HWPF模块是专门用来读取和生成doc格式的文件。在HWPF中，我们使用HWPFDocument来表示一个word doc文档。在看代码之前，有必要了解HWPFDocument中的几个概念：

|名称 | 含义 |  
|:- |:- |
| Range| 表示一个范围，这个范围可以是整个文档，也可以是里面的某个小节（Section），也可以是段落（Paragraph），还可以是拥有功能属性的一段文本（CharacterRun） |
|Section|word文档的一个小节，一个word文档可以由多个小节构成。|
|Paragraph|word文档的一个段落，一个小节可以由多个段落构成。|
|CharacterRun|具有相同属性的一段文本，一个段落可以由多个CharacterRun组成。|
|Table|一个表格。|
|TableRow|表格对应的行|
|TableCell|表格对应的单元格|

**注意 ：  Section、Paragraph、CharacterRun和Table都继承自Range。**


----------

####**读写前注意**：*Apache POI 提供的HWPFDocument类只能读写规范的.doc文件，也就是说假如你使用修改 后缀名 的方式生成doc文件或者直接以命名的方式创建，将会出现错误“Your file appears not to be a valid OLE2 document”*

```
Invalid header signature; read 0x7267617266202E31, expected 0xE11AB1A1E011CFD0 - Your file appears not to be a valid OLE2 document 
```

------------------
####**DOC读**
读doc文件有两种方式
（a）通过WordExtractor读文件
（b）通过HWPFDocument读文件

 在日常应用中，我们从word文件里面读取信息的情况非常少见，更多的还是把内容写入到word文件中。使用POI从word doc文件读取数据时主要有两种方式：通过WordExtractor读和通过HWPFDocument读。在WordExtractor内部进行信息读取时还是通过HWPFDocument来获取的。

#####**使用WordExtractor读**
  在使用WordExtractor读文件时我们只能读到文件的文本内容和基于文档的一些属性，至于文档内容的属性等是无法读到的。如果要读到文档内容的属性则需要使用HWPFDocument来读取了。下面是使用WordExtractor读取文件的一个示例：
```
//通过WordExtractor读文件
public class WordExtractorTest {

   private final String PATH = Environment.getExternalStorageDirectory().getAbsolutePath() + "/" + "test.doc");
   private static final String TAG = "WordExtractorTest";
   
   private void log(Object o) {
       Log.d(TAG, String.valueOf(o));
   }

   public void testReadByExtractor() throws Exception {
      InputStream is = new FileInputStream(PATH);
      WordExtractor extractor = new WordExtractor(is);
      //输出word文档所有的文本
      log(extractor.getText());
      log(extractor.getTextFromPieces());
      //输出页眉的内容
      log("页眉：" + extractor.getHeaderText());
      //输出页脚的内容
      log("页脚：" + extractor.getFooterText());
      //输出当前word文档的元数据信息，包括作者、文档的修改时间等。
      log(extractor.getMetadataTextExtractor().getText());
      //获取各个段落的文本
      String paraTexts[] = extractor.getParagraphText();
      for (int i=0; i<paraTexts.length; i++) {
         log("Paragraph " + (i+1) + " : " + paraTexts[i]);
      }
      //输出当前word的一些信息
      printInfo(extractor.getSummaryInformation());
      //输出当前word的一些信息
      this.printInfo(extractor.getDocSummaryInformation());
      this.closeStream(is);
   }
  
   /**
    * 输出SummaryInfomation
    * @param info
    */
   private void printInfo(SummaryInformation info) {
      //作者
      log(info.getAuthor());
      //字符统计
      log(info.getCharCount());
      //页数
      log(info.getPageCount());
      //标题
      log(info.getTitle());
      //主题
      log(info.getSubject());
   }
  
   /**
    * 输出DocumentSummaryInfomation
    * @param info
    */
   private void printInfo(DocumentSummaryInformation info) {
      //分类
      log(info.getCategory());
      //公司
      log(info.getCompany());
   }
  
   /**
    * 关闭输入流
    * @param is
    */
   private void closeStream(InputStream is) {
      if (is != null) {
         try {
            is.close();
         } catch (IOException e) {
            e.printStackTrace();
         }
      }
   }
}
```

#####**使用HWPFDocument读**
HWPFDocument是当前Word文档的代表，它的功能比WordExtractor要强。通过它我们可以读取文档中的表格、列表等，还可以对文档的内容进行新增、修改和删除操作。只是在进行完这些新增、修改和删除后相关信息是保存在HWPFDocument中的，也就是说我们改变的是HWPFDocument，而不是磁盘上的文件。如果要使这些修改生效的话，我们可以调用HWPFDocument的write方法把修改后的HWPFDocument输出到指定的输出流中。这可以是原文件的输出流，也可以是新文件的输出流（相当于另存为）或其它输出流。下面是一个通过HWPFDocument读文件的示例：
```
//使用HWPFDocument读文件
public class HWPFDocumentTest {
  
   private final String PATH = Environment.getExternalStorageDirectory().getAbsolutePath() + "/" + "test.doc");
   private static final String TAG = "HWPFDocumentTest";
   
   private void log(Object o) {
       Log.d(TAG, String.valueOf(o));
   }

   public void testReadByDoc() throws Exception {
      InputStream is = new FileInputStream(PATH);
      HWPFDocument doc = new HWPFDocument(is);
      //输出书签信息
      this.printInfo(doc.getBookmarks());
      //输出文本
      log(doc.getDocumentText());
      Range range = doc.getRange();
      //读整体
      this.printInfo(range);
      //读表格
      this.readTable(range);
      //读列表
      this.readList(range);
      this.closeStream(is);
   }
  
   /**
    * 关闭输入流
    * @param is
    */
   private void closeStream(InputStream is) {
      if (is != null) {
         try {
            is.close();
         } catch (IOException e) {
            e.printStackTrace();
         }
      }
   }
  
   /**
    * 输出书签信息
    * @param bookmarks
    */
   private void printInfo(Bookmarks bookmarks) {
      int count = bookmarks.getBookmarksCount();
      log("书签数量：" + count);
      Bookmark bookmark;
      for (int i=0; i<count; i++) {
         bookmark = bookmarks.getBookmark(i);
         log("书签" + (i+1) + "的名称是：" + bookmark.getName());
         log("开始位置：" + bookmark.getStart());
         log("结束位置：" + bookmark.getEnd());
      }
   }
  
   /**
    * 读表格
    * 每一个回车符代表一个段落，所以对于表格而言，每一个单元格至少包含一个段落，每行结束都是一个段落。
    * @param range
    */
   private void readTable(Range range) {
      //遍历range范围内的table。
      TableIterator tableIter = new TableIterator(range);
      Table table;
      TableRow row;
      TableCell cell;
      while (tableIter.hasNext()) {
         table = tableIter.next();
         int rowNum = table.numRows();
         for (int j=0; j<rowNum; j++) {
            row = table.getRow(j);
            int cellNum = row.numCells();
            for (int k=0; k<cellNum; k++) {
                cell = row.getCell(k);
                //输出单元格的文本
                log(cell.text().trim());
            }
         }
      }
   }
  
   /**
    * 读列表
    * @param range
    */
   private void readList(Range range) {
      int num = range.numParagraphs();
      Paragraph para;
      for (int i=0; i<num; i++) {
         para = range.getParagraph(i);
         if (para.isInList()) {
            log("list: " + para.text());
         }
      }
   }
  
   /**
    * 输出Range
    * @param range
    */
   private void printInfo(Range range) {
      //获取段落数
      int paraNum = range.numParagraphs();
      log(paraNum);
      for (int i=0; i<paraNum; i++) {
         log("段落" + (i+1) + "：" + range.getParagraph(i).text());
         if (i == (paraNum-1)) {
            this.insertInfo(range.getParagraph(i));
         }
      }
      int secNum = range.numSections();
      log(secNum);
      Section section;
      for (int i=0; i<secNum; i++) {
         section = range.getSection(i);
         log(section.getMarginLeft());
         log(section.getMarginRight());
         log(section.getMarginTop());
         log(section.getMarginBottom());
         log(section.getPageHeight());
         log(section.text());
      }
   }
  
   /**
    * 插入内容到Range，这里只会写到内存中
    * @param range
    */
   private void insertInfo(Range range) {
      range.insertAfter("Hello");
   }
}
```
-------------------
####**DOC写**
#####**使用HWPFDocument写文件**
在使用POI写word doc文件的时候我们必须要先有一个doc文件才行，因为我们在写doc文件的时候是通过HWPFDocument来写的，而HWPFDocument是要依附于一个doc文件的。所以通常的做法是我们先在硬盘上准备好一个内容空白的doc文件，然后建立一个基于该空白文件的HWPFDocument。之后我们就可以往HWPFDocument里面新增内容了，然后再把它写入到另外一个doc文件中，这样就相当于我们使用POI生成了word doc文件。

```
//写字符串进word
    InputStream is = new FileInputStream(PATH);
    HWPFDocument doc = new HWPFDocument(is);

    //获取Range
    Range range = doc.getRange();
    for(int i = 0; i < 100; i++) {
        if( i % 2 == 0 ) {
            range.insertAfter("Hello " + i + "\n");//在文件末尾插入String
        } else {
            range.insertBefore("      Bye " + i + "\n");//在文件头插入String
        }
    }
    //写到原文件中
    OutputStream os = new FileOutputStream(PATH);
    //写到另一个文件中
    //OutputStream os = new FileOutputStream(其他路径);
    doc.write(os);
    this.closeStream(is);
    this.closeStream(os);
```


但是，在实际应用中，我们在生成word文件的时候都是生成某一类文件，该类文件的格式是固定的，只是某些字段不一样罢了。所以在实际应用中，我们大可不必将整个word文件的内容都通过HWPFDocument生成。而是先在磁盘上新建一个word文档，其内容就是我们需要生成的word文件的内容，然后把里面一些属于变量的内容使用类似于“${paramName}”这样的方式代替。这样我们在基于某些信息生成word文件的时候，只需要获取基于该word文件的HWPFDocument，然后调用Range的replaceText()方法把对应的变量替换为对应的值即可，之后再把当前的HWPFDocument写入到新的输出流中。这种方式在实际应用中用的比较多，因为它不但可以减少我们的工作量，还可以让文本的格式更加的清晰。下面我们就来基于这种方式做一个示例。

假设我们有个模板是这样的：
![doc模板](http://upload-images.jianshu.io/upload_images/3971774-fa7da3ca851f9a76.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
之后我们以该文件作为模板，利用相关数据把里面的变量进行替换，然后把替换后的文档输出到另一个doc文件中。具体做法如下：

```
public class HWPFTemplateTest {
	/**
	* 用一个doc文档作为模板，然后替换其中的内容，再写入目标文档中。
	* @throws Exception
	*/
	
	 @Test
   public void testTemplateWrite() throws Exception {
      String templatePath = Environment.getExternalStorageDirectory().getAbsolutePath() + "/" + "template.doc");

      String targetPath = Environment.getExternalStorageDirectory().getAbsolutePath() + "/" + "target.doc";
      InputStream is = new FileInputStream(templatePath);
      HWPFDocument doc = new HWPFDocument(is);
      Range range = doc.getRange();
      //把range范围内的${reportDate}替换为当前的日期
      range.replaceText("${reportDate}", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
      range.replaceText("${appleAmt}", "100.00");
      range.replaceText("${bananaAmt}", "200.00");
      range.replaceText("${totalAmt}", "300.00");
      OutputStream os = new FileOutputStream(targetPath);
      //把doc输出到输出流中
      doc.write(os);
      this.closeStream(os);
      this.closeStream(is);
   }
  
   /**
    * 关闭输入流
    * @param is
    */
   private void closeStream(InputStream is) {
      if (is != null) {
         try {
            is.close();
         } catch (IOException e) {
            e.printStackTrace();
         }
      }
   }
 
   /**
    * 关闭输出流
    * @param os
    */
   private void closeStream(OutputStream os) {
      if (os != null) {
         try {
            os.close();
         } catch (IOException e) {
            e.printStackTrace();
         }
      }
   }
}
```
-------------
### **三、实现docx文件的读写**
POI在读写word docx文件时是通过xwpf模块来进行的，其核心是XWPFDocument。一个XWPFDocument代表一个docx文档，其可以用来读docx文档，也可以用来写docx文档。XWPFDocument中主要包含下面这几种对象：

|对象|含义|
|:---|:---|
|XWPFParagraph|代表一个段落|
|XWPFRun|代表具有相同属性的一段文本|
|XWPFTable|代表一个表格|
|XWPFTableRow|表格的一行|
|XWPFTableCell|表格对应的一个单元格|

####同时XWPFDocument可以直接new一个docx文件出来而不需要像HWPFDocument一样需要一个模板存在。

具体可以参考这位同学写的[POI读写docx文件](http://www.360doc.cn/article/11253639_519415147.html?nsukey=f0Z0Fhcvz%2B8bpq2pzZJfOCgUIRQYCKhRtRXdnELPMGOSm0aD73RHYGQ2JCcEvDljuFY0XJwJ4OTRdA%2F0FKwY4Y07DpT8MKUpvjSFKfrwyiIeSS%2FPmb81efNHXne%2BE5%2FtRtNieEBjFpZzZsO55Y2cKqTQpPwgOkgP8ewdmSeyNO1DIYMUSi6FuKTQj%2BbpJe2S)。

----------------------------------

### **四、总结**
欢迎大家提出建议和纠正本文可能存在的错误之处，感谢支持。
