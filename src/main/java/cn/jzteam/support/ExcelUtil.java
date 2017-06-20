package cn.jzteam.support;

import java.awt.Font;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.sun.rowset.internal.Row;

public class ExcelUtil {

    private static final Logger catalinaLog = Logger.getLogger(ExcelUtil.class);

    private static final String EXCEL_2003 = "xls";

    private static final String EXCEL_2007 = "xlsx";

    /**
     * 从EXCEL表中获取数据信息。
     * 
     * @author liyang E-mail: liyang.bj@liepin.com
     * @version CreateTime：2014年7月25日 下午7:29:35
     * @param multipartFile
     * @param startRowIndex 数据行（标题行不算），从1计数。
     * @return
     * @throws RuntimeException
     * @throws RuntimeException
     */
    public static List<Map<String, Object>> processDataFormExcel(MultipartFile multipartFile, int startRowIndex)
            throws RuntimeException {
        // 校验文件信息
        Map<String, String> result = validateFile(multipartFile);
        if ("0".equals(result.get("code"))) {
            throw new RuntimeException(result.get("msg")); // 操作异常
        } else if ("1".equals(result.get("code"))) {
            throw new RuntimeException(result.get("msg")); // 系统异常
        }
        List<Map<String, Object>> dataList = new ArrayList<Map<String, Object>>();
        // 获取文件名称
        String fileName = multipartFile.getOriginalFilename();

        InputStream input = new ByteArrayInputStream(multipartFile.getBytes());
        // 声明一个工作簿
        Workbook workbook = null;
        if (fileName.endsWith(EXCEL_2003)) { // 2003版本的EXCEL文件
            workbook = new HSSFWorkbook(input);
        } else if (fileName.endsWith(EXCEL_2007)) { // 2007版本的EXCEL文件
            workbook = new XSSFWorkbook(input);
        } else {
            throw new RuntimeException("文件格式不正确");
        }
        // 获取当前Sheet页的工作表【只支持一个Sheet页】
        Sheet sheet = workbook.getSheetAt(0);
        // 获取工作表的总行数
        int rowCount = sheet.getPhysicalNumberOfRows();
        if (startRowIndex > rowCount) {
            throw new RuntimeException("数据的起始行数超过了总行数，请检查！");
        }
        // 获取标题行
        Row titleRow = sheet.getRow(0);
        if (titleRow == null) {
            throw new RuntimeException("标题行为空，请检查！");
        }

        // 获取总列数
        int colCount = titleRow.getLastCellNum();
        // 从第startRowIndex行开始，第1行是标题栏
        for (int rowIndex = startRowIndex; rowIndex < rowCount; rowIndex++) {
            try {
                Row currentRow = sheet.getRow(rowIndex);
                // 定义返回的数据：key：列数，从1 开始；object ：列对应的值 一个Map，一行数据
                Map<String, Object> mapValue = new HashMap<String, Object>();
                // 读取每一列的信息
                for (int colIndex = 0; colIndex < colCount; colIndex++) {
                    Cell cell = currentRow.getCell(colIndex);
                    String cellValue = null;
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_NUMERIC: // 数字
                                // 先看是否是日期格式
                                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                    // 读取日期格式
                                    Date dateValue = cell.getDateCellValue();
                                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-mm-dd");
                                    cellValue = sdf.format(dateValue);
                                } else {
                                    cell.setCellType(Cell.CELL_TYPE_STRING);
                                    cellValue = cell.getRichStringCellValue().toString();
                                }
                                break;
                            case Cell.CELL_TYPE_STRING: // 字符串
                                if (!"".equals(cell.getStringCellValue().trim())) {
                                    cellValue = cell.getStringCellValue();
                                }
                                break;
                            case Cell.CELL_TYPE_BOOLEAN: // Boolean
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case Cell.CELL_TYPE_FORMULA: // 公式
                                cellValue = String.valueOf(cell.getCellFormula());
                                break;
                            case Cell.CELL_TYPE_BLANK: // 空值
                                break;
                            default:
                                break;
                        }
                        if (cellValue != null) {
                            mapValue.put(String.valueOf(colIndex + 1), cellValue);
                        }
                    }
                }
                if (!mapValue.isEmpty()) {
                    dataList.add(mapValue);
                }
            } catch (RuntimeException e) {
                catalinaLog.info(e.getMessage());
            }
        }
        return dataList;
    }

    /**
     * 将数据导出到Excel 2007版本
     * 
     * @author liyang E-mail: liyang.bj@liepin.com
     * @version CreateTime：2014年10月9日 下午2:22:22
     * @param title :表格的标题名称
     * @param headers ：表格属性列名称的数组
     * @param dataMap ：需要导出的数据
     * @param OutputStream ：导出的Excel文件存储的位置
     * @throws IORuntimeException
     * @throws InvocationTargetRuntimeException
     * @throws IllegalAccessRuntimeException
     * @throws IllegalArgumentRuntimeException
     */
    public static void export2007ExcelByPOI(String title, String[] headers, List<Map<String, Object>> dataMap,
            OutputStream outputStream) throws IORuntimeException, IllegalArgumentRuntimeException,
            IllegalAccessRuntimeException, InvocationTargetRuntimeException {
        // 声明一个工作簿【SXSSFWorkbook只支持.xlsx格式】
        Workbook workbook = new SXSSFWorkbook(1000);// 内存中只存放1000条
        // 生成一个表格
        Sheet sheet = workbook.createSheet(title);
        // 设置表格的默认宽度为18个字节
        sheet.setDefaultColumnWidth(18);
        // 生成一个样式【用于表格标题】
        CellStyle headStyle = workbook.createCellStyle();
        // 设置样式
        headStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        headStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        headStyle.setBorderTop(CellStyle.BORDER_THIN);// 单元格上边框
        headStyle.setBorderBottom(CellStyle.BORDER_THIN);// 单元格下边框
        headStyle.setBorderLeft(CellStyle.BORDER_THIN);// 单元格左边框
        headStyle.setBorderRight(CellStyle.BORDER_THIN);// 单元格右边框
        headStyle.setAlignment(CellStyle.ALIGN_CENTER);// 单元格水平居中
        headStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 单元格垂直居中
        // 生成字体【用于表格标题】
        Font headFont = workbook.createFont();
        headFont.setFontHeightInPoints((short) 12);
        headFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        // 把字体应用到当前样式
        headStyle.setFont(headFont);

        // 生成一个样式【用于Excel中的表格内容】
        CellStyle contentStyle = workbook.createCellStyle();
        // 设置样式【用于Excel中的表格内容】
        contentStyle.setBorderTop(CellStyle.BORDER_THIN);// 单元格上边框
        contentStyle.setBorderBottom(CellStyle.BORDER_THIN);// 单元格下边框
        contentStyle.setBorderLeft(CellStyle.BORDER_THIN);// 单元格左边框
        contentStyle.setBorderRight(CellStyle.BORDER_THIN);// 单元格右边框
        contentStyle.setAlignment(CellStyle.ALIGN_CENTER);// 单元格水平居中
        contentStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 单元格垂直居中
        contentStyle.setWrapText(true);// 单元格自动换行
        // 生成字体
        Font contentFont = workbook.createFont();
        contentFont.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前样式
        contentStyle.setFont(contentFont);

        // 产生表格标题行【表格的第一行】
        Row headRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headRow.createCell(i);
            // 设置单元格为文本格式
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellStyle(headStyle);
            RichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        // 遍历集合数据，产生EXCEL行【Excel表格的标题占用了一行】
        int index = 1;
        for (Map<String, Object> temp : dataMap) {
            // 创建一行
            Row row = sheet.createRow(index);
            Set<String> keys = temp.keySet();
            for (int i = 0; i < keys.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(contentStyle);
                Object value = temp.get(String.valueOf(i));
                String textValue = "";
                // 判断类型之后进行类型转换
                if (value == null) {
                    cell.setCellValue(textValue);
                } else if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                } else if (value instanceof Long) {
                    cell.setCellValue((Long) value);
                } else if (value instanceof Double) {
                    cell.setCellValue((Double) value);
                } else if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                } else if (value instanceof Date) {
                    DateFormat defaultFormatter = new SimpleDateFormat("yyyy-MM-dd");
                    textValue = defaultFormatter.format((Date) value);
                } else if (value instanceof byte[]) {
                    byte[] pictureData = (byte[]) value;
                    // 有图片时候，设置行高
                    row.setHeightInPoints(60);
                    // 设置图片所在的列为80px
                    sheet.setColumnWidth(i, (short) (35.7 * 80));
                    XSSFClientAnchor anchorOher = new XSSFClientAnchor(0, 0, 1023, 255, (short) 6, index, (short) 6,
                            index);
                    anchorOher.setAnchorType(2);
                    // 声明一个画图的顶级管理器
                    Drawing patriarch = sheet.createDrawingPatriarch();
                    patriarch.createPicture(anchorOher,
                            workbook.addPicture(pictureData, HSSFWorkbook.PICTURE_TYPE_JPEG));
                } else {
                    textValue = value.toString();
                }
                // 如果是字符串
                if (!"".equals(textValue)) {
                    RichTextString richString = new XSSFRichTextString(textValue);
                    // 设置单元格为文本格式
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    cell.setCellValue(richString);
                }
            }
            // 行加1
            index++;
        }
        workbook.write(outputStream);
        outputStream.flush();
    }

    /**
     * 获取一个格里的值
     * 
     * @param row
     * @param cellIndex
     * @return
     */
    public static String getCellValue(Row row, int cellIndex) {
        String value = "";
        Cell cell = row.getCell(cellIndex);
        if (cell == null) {
            return value;
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                // 得到Boolean对象的方法
                value = cell.getBooleanCellValue() + "";
                break;
            case Cell.CELL_TYPE_NUMERIC:
                // 先看是否是日期格式
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    // 读取日期格式
                    Date dateValue = cell.getDateCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
                    value = sdf.format(dateValue);
                } else {
                    cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                    value = cell.getRichStringCellValue().toString();
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                // 读取公式
                value = cell.getCellFormula();
                break;
            case Cell.CELL_TYPE_STRING:
                // 读取String
                value = cell.getRichStringCellValue().toString();
                break;
        }
        return value;
    }

    /**
     * 校验文件基本信息（文件名，大小）
     * 
     * @author liyang E-mail: liyang.bj@liepin.com
     * @version CreateTime：2014年7月25日 下午7:22:25
     * @param multipartFile
     * @return
     */
    public static Map<String, String> validateFile(MultipartFile multipartFile) {
        Map<String, String> result = new HashMap<String, String>();
        try {
            // 默认是文件正常
            result.put("code", "2");
            result.put("msg", "文件校验正常！");
            // 获取文件名称
            String fileName = multipartFile.getOriginalFilename();
            // 校验文件的类型
            if (!fileName.endsWith(EXCEL_2003) && !fileName.endsWith(EXCEL_2007)) {
                result.put("code", "0");
                result.put("msg", "请选择后缀类型为.xlsx的Excel文件！");
            } else {
                // 校验文件的大小
                // 获取上传文件的大小【fis.available()单位是byte】
                double dataSize = (double) multipartFile.getSize() / (1024 * 1024);
                if (Double.compare(dataSize, 0D) == 0) {
                    result.put("code", "0");
                    result.put("msg", "请选择需要上传的文件！");
                }
                if (Double.compare(dataSize, 10) > 0) {
                    result.put("code", "0");
                    result.put("msg", "上传的文件超过10M,请拆分后上传！");
                }
            }
        } catch (RuntimeException e) {
            result.put("code", "1");
            result.put("msg", "系统出现错误，请稍后再试！");
        }
        return result;
    }

    /**
     * 
     * 类描述：参数验证类型
     * 
     * @author: lihonghao
     * @date： 日期：2014-11-7 时间：下午07:35:45
     * @version 1.0
     */
    public enum EnumCellType {
        STRING, PHONE, DATE, DQ, INDUSTRY
    }
}
