package com.king.zxing.app;
import android.os.Build;

import androidx.annotation.RequiresApi;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * @author LiHangZhou
 * @description:
 * @date :2021/1/6 10:07
 * @love :zlx
 */


public class ExcelDataUtil {

    @RequiresApi(api = Build.VERSION_CODES.N)
    public static void main(String[] args) {
        try {
            List<Map<String, String>> maps = redExcel("D:\\文档\\cp方对接参数.xlsx");
            System.out.println(maps);
        } catch (Exception e) {
            System.out.println(e);
        }
//        写入
//        File file = new File("D:\\文档\\a.xlsx");
//        WriteInExcel(file,"ad",0);

//        //打开文件
//        WritableWorkbook book = Workbook.createWorkbook(new File("D:\\文档\\a.xlsx"));
//        //生成名为“第一页”的工作表，参数0表示这是第一页
//        WritableSheet sheet = book.createSheet("第一页", 0);
//        //在Label对象的构造子中指名单元格位置是第一列第一行(0,0)
//        //以及单元格内容为test
//        Label label = new Label(0, 0, "测试");
//        //将定义好的单元格添加到工作表中
//        sheet.addCell(label);
//        jxl.write.Number number = new jxl.write.Number(1, 0, 789.123);
//        sheet.addCell(number);
//        jxl.write.Label s = new jxl.write.Label(1, 2, "三十三");
//        sheet.addCell(s);
//        //写入数据并关闭文件
//        book.write();
//        book.close(); //最好在finally中关闭，此处仅作为示例不太规范
    }

    /**
     *
     * @param file 文件路径
     * @param pagination 第几页
     * @param i 参数0表示这是第一页
     * @throws IOException
     * @throws WriteException
     */
    private static void WriteInExcel(File file, String pagination, int i) throws IOException, WriteException {
        //打开文件
        WritableWorkbook book = Workbook.createWorkbook(file);
        //生成名为“第一页”的工作表，参数0表示这是第一页
        WritableSheet sheet = book.createSheet(pagination, i);
        //在Label对象的构造子中指名单元格位置是第一列第一行(0,0)
        //以及单元格内容为test
        Label label = new Label(0, 0, "测试");
        //将定义好的单元格添加到工作表中
        sheet.addCell(label);
        jxl.write.Number number = new jxl.write.Number(1, 0, 789.123);
        sheet.addCell(number);
        jxl.write.Label s = new jxl.write.Label(1, 2, "三十三");
        sheet.addCell(s);
        System.out.println("写入成功");
        //写入数据并关闭文件
        book.write();
        book.close(); //最好在finally中关闭，此处仅作为示例不太规范
    }


    /**
     * 读取excel内容
     * <p>
     * 用户模式下：
     * 弊端：对于少量的数据可以，单数对于大量的数据，会造成内存占据过大，有时候会造成内存溢出
     * 建议修改成事件模式
     */
    public static List<Map<String, String>> redExcel(String filePath) throws Exception {
        File file = new File(filePath);
        if (!file.exists()) {
            throw new Exception("文件不存在!");
        }
        InputStream in = new FileInputStream(file);

        // 读取整个Excel
        XSSFWorkbook sheets = new XSSFWorkbook(in);
        // 获取第一个表单Sheet
        XSSFSheet sheetAt = sheets.getSheetAt(0);
        ArrayList<Map<String, String>> list = new ArrayList<>();

        //默认第一行为标题行，i = 0
        XSSFRow titleRow = sheetAt.getRow(0);
        // 循环获取每一行数据
        for (int i = 1; i < sheetAt.getPhysicalNumberOfRows(); i++) {
            XSSFRow row = sheetAt.getRow(i);
            LinkedHashMap<String, String> map = new LinkedHashMap<>();
            // 读取每一格内容
            for (int index = 0; index < row.getPhysicalNumberOfCells(); index++) {
                XSSFCell titleCell = titleRow.getCell(index);
                XSSFCell cell = row.getCell(index);
                // cell.setCellType(XSSFCell.CELL_TYPE_STRING); 过期，使用下面替换
                cell.setCellType(CellType.STRING);
                if (cell.getStringCellValue().equals("")) {
                    continue;
                }
                map.put(getString(titleCell), getString(cell));
            }
            if (map.isEmpty()) {
                continue;
            }
            list.add(map);
        }
        return list;
    }

    /**
     * 把单元格的内容转为字符串
     *
     * @param xssfCell 单元格
     * @return String
     */
    public static String getString(XSSFCell xssfCell) {
        if (xssfCell == null) {
            return "";
        }
        if (xssfCell.getCellTypeEnum() == CellType.NUMERIC) {
            return String.valueOf(xssfCell.getNumericCellValue());
        } else if (xssfCell.getCellTypeEnum() == CellType.BOOLEAN) {
            return String.valueOf(xssfCell.getBooleanCellValue());
        } else {
            return xssfCell.getStringCellValue();
        }
    }
}