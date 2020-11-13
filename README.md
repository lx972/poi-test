# POI报表

## 1 依赖

目前的最新版本

```xml
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>4.1.2</version>
</dependency>

<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.1.2</version>
</dependency>

<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml-schemas -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml-schemas</artifactId>
    <version>4.1.2</version>
</dependency>
```

## 2 poi结构

**HSSF**提供读写**Microsoft Excel XLS**格式档案的功能。

**XSSF**提供读写**Microsoft Excel OOXML XLSX**格式档案的功能。

**HWPF**提供读写**Microsoft Word DOC**格式档案的功能。

**HSLF**提供读写**Microsoft PowerPoint**格式档案的功能。

**HDGF**提供读**Microsoft Visio**格式档案的功能。

**HPBF**提供读**Microsoft Publisher**格式档案的功能。

**HSMF**提供读**Microsoft Outlook**格式档案的功能。

> 根据这个结构创建对应的工作空间，即可操作对应格式的文档。

## 3 xssf表格

### 3.1 创建简单的表格

```java
/**
 * 简单的创建一个表格
 */
@Test
public void test1() {
    //创建一个工作空间（表格xlsx）
    XSSFWorkbook workbook = new XSSFWorkbook();
    //创建一个sheet
    XSSFSheet sheet1 = workbook.createSheet("test1");
    XSSFSheet sheet2 = workbook.createSheet("test2");

     //获取sheet中某一行
    XSSFRow row1 = sheet1.createRow(3);
    //获取一行中的某个单元格
    XSSFCell cell1 = row1.createCell(3);
     //设置该单元格的内容
    cell1.setCellValue(43543);

    //这个和上面差不多，只不过是操作另一个sheet
    XSSFRow row2 = sheet2.createRow(5);
    XSSFCell cell2 = row2.createCell(3);
    cell2.setCellValue("dfgdfhgfdhfghfg");

    FileOutputStream os = null;
    try {
        //写到磁盘中
        os = new FileOutputStream("D:\\BaiduNetdiskDownload\\test1.xlsx");
        workbook.write(os);
        os.close();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }
}
```

### 3.2 创建有样式的表格

```java
/**
     * 设置单元格样式
     */
@Test
public void test2() {
    //创建一个工作空间（表格xlsx）
    XSSFWorkbook workbook = new XSSFWorkbook();
    //创建一个sheet
    XSSFSheet sheet = workbook.createSheet("test1");

    XSSFRow row1 = sheet.createRow(3);
    XSSFCell cell1 = row1.createCell(3);
    //设置单元格值类型
    cell1.setCellType(CellType.STRING);
    //设置单元格内容
    cell1.setCellValue(43543);    
    
    //设置单元格样式
    //单元格样式要从工作空间创建，不能直接new
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    //设置单元格顶部线条的粗细程度
    cellStyle.setBorderTop(BorderStyle.THIN);
    //设置单元格顶部边框颜色为红色
    cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, new XSSFColor(Color.RED));
    cell1.setCellStyle(cellStyle);
    
    //设置字体
    XSSFFont font = workbook.createFont();
    font.setColor(IndexedColors.LEMON_CHIFFON.getIndex());
    cellStyle.setFont(font);
    
    FileOutputStream os = null;
    try {
        //写到磁盘中
        os = new FileOutputStream("D:\\BaiduNetdiskDownload\\test2.xlsx");
        workbook.write(os);
        os.close();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }
}
```

### 3.3 在表格中绘制图片

```java
/**
 * 在表格中绘制图片
 */
@Test
public void test3() {
    //创建一个工作空间（表格xlsx）
    XSSFWorkbook workbook = new XSSFWorkbook();
    //创建一个sheet
    XSSFSheet sheet = workbook.createSheet("test1");

    try {
        //将读片读到输入流中
        FileInputStream is = new FileInputStream("C:\\Users\\Administrator\\Desktop\\lua\\image\\5.png");
        //将输入流中的图片读到工作空间中，需要指定图片格式
        int pictureIndex = workbook.addPicture(is, Workbook.PICTURE_TYPE_PNG);

        //获取表格的帮助类
        XSSFCreationHelper creationHelper = workbook.getCreationHelper();
        //由帮助类创建一个锚点
        XSSFClientAnchor clientAnchor = creationHelper.createClientAnchor();
        //设置锚点的起始坐标（按单元格计量）
        clientAnchor.setCol1(4);
        clientAnchor.setRow1(0);
        //创建一个绘图对象
        XSSFDrawing drawingPatriarch = sheet.createDrawingPatriarch();
        //绘制（需要指定绘制的起始坐标，和图片在工作空间上的索引）
        XSSFPicture picture = drawingPatriarch.createPicture(clientAnchor, pictureIndex);
        //将图片的尺寸重置为嵌入的尺寸
        picture.resize();

        //写到磁盘中
        FileOutputStream os = new FileOutputStream("D:\\BaiduNetdiskDownload\\test2.xlsx");
        workbook.write(os);
        os.close();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }
}

```

### 3.4 读取磁盘中表格内容

```java
/**
 * 读取表格信息
 */
@Test
public void test4() {
    //声明一个工作空间（表格xlsx）
    XSSFWorkbook workbook = null;
    try {
        //将表格信息读取到该工作空间中
        workbook = new XSSFWorkbook("D:\\BaiduNetdiskDownload\\test2.xlsx");
    } catch (IOException e) {
        e.printStackTrace();
    }
    //获取一个sheet
    XSSFSheet sheet = workbook.getSheetAt(0);

    //循环所有行
    for (int i = 0; i < sheet.getLastRowNum(); i++) {
        //具体某一行
        XSSFRow row = sheet.getRow(i);
        //该行为空直接读取下一行
        if (row==null){
            continue;
        }
        //循环该行所有单元格
        for (int j = 0; j < row.getLastCellNum(); j++) {
            XSSFCell cell = row.getCell(j);
            //该单元格为空，直接读取下一个单元格
            if (cell==null){
                continue;
            }
            //判断单元格值的类型，然后取值
            switch (cell.getCellType()) {
                case STRING:
                    System.out.println("string:"+cell.getStringCellValue());
                    break;
                case BOOLEAN:
                    System.out.println("boolean:"+cell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    //日期格式比较麻烦，它在报表中存储的是数字类型，所以必须先使用自带的工具类判断
                    if ((DateUtil.isCellDateFormatted(cell))) {
                        System.out.println("日期:" + cell.getDateCellValue());
                    } else {
                        System.out.println("number:" + cell.getNumericCellValue());
                    }
                    break;
            }
        }
    }
}
```

### 3.5 浏览器下载报表

要实现浏览器下载功能，必须设置请求头，告诉浏览器此次访问是下载，浏览器才会使用对应的下载功能。

对于火狐浏览器

```java
String fileName = URLEncoder.encode("员工信息表.xlsx","UTF-8");
response.setContentType("application/octet-stream");
//注意，构建的文件名必须是ISO8859-1格式的，utf-8不行
response.addHeader("content-disposition",
                   "attachment;filename=" +
                   new String(fileName.getBytes(Charset.defaultCharset()),"ISO8859-1"));
ServletOutputStream outputStream = response.getOutputStream();
workbook.write(outputStream);
```

### 3.6 模板报表

使用代码控制格式会影响代码的阅读，而且项目上线以后，格式难以修改。所以我们一般设置一个模板表格文件，在这个表格文件中，我们先将表头和格式设置好。在项目中直接读取该模板文件，然后填充数据返回给浏览器下载。

```java
//获取根路径下的模板文件
ClassPathResource classPathResource = new ClassPathResource("templates.xlsx");

XSSFWorkbook workbook = null;
try {
    //将模板文件读取到工作空间中
    workbook = new XSSFWorkbook(classPathResource.getInputStream());
} catch (IOException e) {
    e.printStackTrace();
}
XSSFSheet sheet = workbook.getSheetAt(0);
//获取第一行表头，所有格式均遵从第一行
XSSFRow row = sheet.getRow(0);

//单元格样式，一列一个样式
XSSFCellStyle[] cellStyle=new XSSFCellStyle[row.getLastCellNum()+1];
for (int i = 0; i < row.getLastCellNum(); i++) {
    cellStyle[i] = row.getCell(i).getCellStyle();
}

//表内容
User user = null;
for (int i =0; i < userList.size(); i++) {
    XSSFRow content = sheet.createRow(i+1);
    user = userList.get(i);
    
    XSSFCell cell0 = content.createCell(0);
    cell0.setCellValue(user.getUsername());
    cell0.setCellStyle(cellStyle[0]);

    XSSFCell cell1 = content.createCell(1);
    cell1.setCellValue(user.getMobile());
    cell1.setCellStyle(cellStyle[1]);

    XSSFCell cell2 = content.createCell(2);
    cell2.setCellValue(user.getWorkNumber());
    cell2.setCellStyle(cellStyle[2]);

    XSSFCell cell3 = content.createCell(3);
    cell3.setCellValue(user.getFormOfEmployment());
    cell3.setCellStyle(cellStyle[3]);

    XSSFCell cell4 = content.createCell(4);
    cell4.setCellValue(new SimpleDateFormat("yyyy-MM-dd").format(user.getTimeOfEntry()));
    cell4.setCellStyle(cellStyle[4]);

}
try {
    String fileName = URLEncoder.encode(month + "月员工信息表.xlsx",
                                        "UTF-8");
    response.setContentType("application/octet-stream");
    response.addHeader("content-disposition",
                       "attachment;filename=" +
                       new String(fileName.getBytes(Charset.defaultCharset()),"ISO8859-1"));
    response.addHeader("fileName", fileName);
    ServletOutputStream outputStream = response.getOutputStream();
    workbook.write(outputStream);

} catch (Exception e) {
    log.error(e.getMessage(), e);
    throw new CommonException(ResultCode.SERVER_ERROR);
}
```

### 3.7 百万数据报表导出

基于`XSSFWorkbook`的报表是将所有数据加载到内存中，在内存中操作，当数据量达到百万级别，内存中根本无法存下这么多的数据，为了解决这个问题，Apache Poi提供了`SXSSFWork`对象，专门用于处理大数据量Excel报表导出。

```java
/**
 * 大数据量专用的工作空间
 * 原理很简单，就是每次数据量达到100条数据的时候，就将数据写到磁盘中，同时
 * 将这些数据从内存中销毁，避免内存溢出
 */
@Test
public void test1(){

    //创建一个工作空间（表格xlsx）
    SXSSFWorkbook workbook = new SXSSFWorkbook();

    SXSSFSheet sheet = workbook.createSheet();
    for (int i = 0; i < 1000000; i++) {
        SXSSFRow row = sheet.createRow(i);
        for (int j = 0; j < 6; j++) {
            SXSSFCell cell = row.createCell(j);
            cell.setCellValue(i+j+"中国");
        }
    }
    FileOutputStream os = null;
    try {
        //写到磁盘中
        os = new FileOutputStream("D:\\BaiduNetdiskDownload\\test1.xlsx");
        workbook.write(os);
        os.close();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }
}


/**
 * 小数据量可以用这个
 * 对比运行这两个方法，可以明显发现test2方法需要的时间特别长
 * 还有可能触发异常java.lang.OutOfMemoryError: GC overhead limit exceeded
 */
@Test
public void test2(){
    //创建一个工作空间（表格xlsx）
    XSSFWorkbook workbook = new XSSFWorkbook();

    XSSFSheet sheet = workbook.createSheet();
    for (int i = 0; i < 1000000; i++) {
        XSSFRow row = sheet.createRow(i);
        for (int j = 0; j < 6; j++) {
            XSSFCell cell = row.createCell(j);
            cell.setCellValue(i+j+"中国");
        }
    }
    FileOutputStream os = null;
    try {
        //写到磁盘中
        os = new FileOutputStream("D:\\BaiduNetdiskDownload\\test2.xlsx");
        workbook.write(os);
        os.close();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }
}
```

### 3.8 百万数据报表读取

读取的时候，有两种模式：

（1）用户模式：一次性将所有数据加载到内存中，然后再对单元格一个一个解析。这种模式不用说，在数据量大的时候，很容易造成内存泄漏。

（2）事件模式：逐行扫描，然后解析（实际上就是一边扫描一边解析）。

```java
/**
 * 事件模式
 */
@Test
public void test() throws Exception {
    OPCPackage opcPackage = null;
    try {
        //创建一个可以存储多个数据对象的容器
        opcPackage = OPCPackage.open("D:\\BaiduNetdiskDownload\\test1.xlsx", PackageAccess.READ);

        //根据该容器创建一个表格读取器
        XSSFReader reader = new XSSFReader(opcPackage);
        //获取表共享字符串对象
        SharedStringsTable sharedStringsTable = reader.getSharedStringsTable();
        //获取表样式对象
        StylesTable stylesTable = reader.getStylesTable();
        //使用xml读取器工厂创建一个解析器
        XMLReader xmlReader = XMLReaderFactory.createXMLReader();

        //注册一个事件处理器，并在事件处理器中设置对解析后的内容如何操作
        xmlReader.setContentHandler(new XSSFSheetXMLHandler(stylesTable, sharedStringsTable, new XSSFSheetXMLHandler.SheetContentsHandler() {
            /**
                 * A row with the (zero based) row number has started
                 *
                 * @param rowNum
                 */
            public void startRow(int rowNum) {
                log.info("开始解析第" + rowNum + "行");
            }

            /**
                 * A row with the (zero based) row number has ended
                 *
                 * @param rowNum
                 */
            public void endRow(int rowNum) {
                log.info("完成解析第" + rowNum + "行");
            }

            /**
                 * A cell, with the given formatted value (may be null),
                 * and possibly a comment (may be null), was encountered.
                 * <p>
                 * Sheets that have missing or empty cells may result in
                 * sparse calls to <code>cell</code>. See the code in
                 * <code>src/examples/src/org/apache/poi/xssf/eventusermodel/XLSX2CSV.java</code>
                 * for an example of how to handle this scenario.
                 *
                 * @param cellReference 列标识，就是表格列的A,B,C,D
                 * @param formattedValue 单元格中的值
                 * @param comment 单元格注释
                 */
            public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                log.info("cellReference：{}，formattedValue：{}", cellReference, formattedValue);
            }
        }, false));

        //获取数据流的迭代器，开始一行一行读取数据
        Iterator<InputStream> sheetsData = reader.getSheetsData();
        while (sheetsData.hasNext()) {
            //获取改行数据的字节奎流
            InputStream inputStream = sheetsData.next();
            //使用字节流创建一个新的输入源
            InputSource inputSource = new InputSource(inputStream);
            try {
                //解析输入源中的内容
                xmlReader.parse(inputSource);
            } finally {
                inputStream.close();
            }
        }
    } finally {
        opcPackage.close();
    }
}
```

