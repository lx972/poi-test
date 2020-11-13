import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.junit.Test;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * PACKAGE_NAME
 *
 * @Author Administrator
 * @date 15:48
 */
public class XSSF {

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

        XSSFRow row1 = sheet1.createRow(3);
        XSSFCell cell1 = row1.createCell(3);
        cell1.setCellValue(43543);

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
        cell1.setCellValue(43543);
        //设置单元格样式
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
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

    /**
     * 读取表格信息
     */
    @Test
    public void test4() {
        //创建一个工作空间（表格xlsx）
        XSSFWorkbook workbook = null;
        try {
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
            if (row==null){
                continue;
            }
            //循环该行所有单元格
            for (int j = 0; j < row.getLastCellNum(); j++) {
                XSSFCell cell = row.getCell(j);
                if (cell==null){
                    continue;
                }
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println("string:"+cell.getStringCellValue());
                        break;
                    case BOOLEAN:
                        System.out.println("boolean:"+cell.getBooleanCellValue());
                        break;
                    case NUMERIC:
                        System.out.println("number:"+cell.getNumericCellValue());
                        break;
                }
            }

        }
    }
}
