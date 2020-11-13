import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

/**
 * PACKAGE_NAME
 *
 * @Author Administrator
 * @date 15:44
 */
@Slf4j
public class SXSSF {


    /**
     * 大数据量专用的工作空间
     * 原理很简单，就是每次数据量达到100条数据的时候，就将数据写到磁盘中，同时
     * 将这些数据从内存中销毁，避免内存溢出
     */
    @Test
    public void test1() {

        //创建一个工作空间（表格xlsx）
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        SXSSFSheet sheet = workbook.createSheet();
        for (int i = 0; i < 1000000; i++) {
            SXSSFRow row = sheet.createRow(i);
            for (int j = 0; j < 6; j++) {
                SXSSFCell cell = row.createCell(j);
                cell.setCellValue(i + j + "中国");
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
    public void test2() {
        //创建一个工作空间（表格xlsx）
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet();
        for (int i = 0; i < 1000000; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < 6; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(i + j + "中国");
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


    @Test
    public void test3() throws Exception {
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


}
