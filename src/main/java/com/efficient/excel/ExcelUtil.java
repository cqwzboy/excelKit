package com.efficient.excel;


import com.efficient.excel.annotation.DateFormat;
import com.efficient.excel.annotation.ExcelHeader;
import com.efficient.excel.annotation.ExcelTypeHandler;
import com.efficient.excel.domain.ExcelHeaderWrapper;
import com.efficient.excel.enums.CellStylePosition;
import com.efficient.excel.enums.ExcelType;
import com.qc.itaojin.enums.GenericType;
import com.qc.itaojin.util.ReflectUtils;
import com.qc.itaojin.util.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

/**
 * 
 * excel工具类
 *
 * @author qinqin Fu
 * @since 2018-10-11
 */
@Slf4j
public final class ExcelUtil {
    /**
     * 默认sheet名称
     * */
    public static final String DEFAULT_SHEETNAME = "sheet1";

    /**
     * 导入
     * @param path 文件所在目录
     * @param sheetName sheet名称
     * @param clazz 返回数据类型
     * */
    public static <T> List<T> importData(String path, String sheetName, Class<T> clazz){
        if(StringUtils.isBlank(path) || clazz==null || StringUtils.isBlank(sheetName)){
            throw new IllegalArgumentException("error parameters");
        }

        List<T> list = new ArrayList<>();

        // 解析Wrokbook
        Workbook wb = parseWorkbook(path);

        // 解析sheet
        Sheet sheet = parseSheet(wb, sheetName);
        if(isEmpty(sheet)){
            return list;
        }

        // Short-column下角标 String-Excel标题
        Map<Short, String> titleMap = new HashMap<>();
        // 解析标题
        Row headerRow = sheet.getRow(0);
        short firstCellNum = headerRow.getFirstCellNum();
        short lastCellNum = headerRow.getLastCellNum();
        for(short i=firstCellNum;i<lastCellNum;i++){
            titleMap.put(i, headerRow.getCell(i).getStringCellValue());
        }

        // 解析实体类clazz，解析出标题名和属性的对应关系 String-标题 Field-属性
        Map<String, Field> titleAndTypeMap = new HashMap<>();
        for (Field field : clazz.getDeclaredFields()) {
            // 如果属性没有注解ExcelHeaderColumn，跳过
            if(!ReflectUtils.hasAnnotationPresent(field, ExcelHeader.class)){
                continue;
            }
            String title = ReflectUtils.analyzeFieldAnnotation(field, ExcelHeader.class, "title");
            titleAndTypeMap.put(title, field);
        }

        // 计算column下角标和属性名的对应关系 Short-column下角标 Field-属性
        Map<Short, Field> fieldMap = new HashMap<>();
        for (Short key : titleMap.keySet()) {
            fieldMap.put(key, titleAndTypeMap.get(titleMap.get(key)));
        }

        // 解析主体内容
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getPhysicalNumberOfRows();
        out:for(int i=firstRowNum+1;i<lastRowNum;i++){
            Row row = sheet.getRow(i);
            T t;
            try {
                t = clazz.newInstance();
            } catch (Exception e) {
                throw new RuntimeException(e.getMessage());
            }
            in:for(short j=firstCellNum;j<lastCellNum;j++){
                // 当Cell不存在时返回一个空的cell
                Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                // 获取单元格值类型
                CellType cellType = cell.getCellTypeEnum();
                Field field = fieldMap.get(j);
                Object param = null;
                switch (cellType){
                    case STRING:
                        if(ReflectUtils.hasAnnotationPresent(field, ExcelTypeHandler.class)){
                            Class typeHandlerClazz = ReflectUtils.analyzeFieldAnnotation(field, ExcelTypeHandler.class, "clazz");
                            try {
                                param = ((com.efficient.excel.enums.ExcelTypeHandler)typeHandlerClazz.newInstance()).onImport(cell.getStringCellValue());
                            } catch (Exception e) {
                                throw new RuntimeException("error on parse @ExcelTypeHandler");
                            }
                        }else{
                            param = ReflectUtils.transValue(field, cell.getStringCellValue());
                        }
                        break;
                    case NUMERIC:
                        if(DateUtil.isCellDateFormatted(cell)){// 日期
                            param = cell.getDateCellValue();
                        }else{// 数字型
                            String numericString = String.valueOf(cell.getNumericCellValue());
                            if(!isDecimal(cell.getNumericCellValue())){
                                numericString = numericString.substring(0, numericString.indexOf("."));
                            }
                            param = ReflectUtils.transValue(field, numericString);
                        }
                        break;
                    case BOOLEAN:
                        param = cell.getBooleanCellValue();
                        break;
                    case BLANK:
                        continue in;
                    default:
                        continue in;
                }
                ReflectUtils.invokeSet(t, field.getName(), new Object[]{param});
            }
            list.add(t);
        }

        return list;
    }

    /**
     * 导出
     * @param path 文件所在目录
     * @param sheetName sheet名称
     * @param list 源数据
     * */
    public static <T> void exportData(String path, String sheetName, List<T> list){
        if(StringUtils.isBlank(path) || CollectionUtils.isEmpty(list)){
            throw new IllegalArgumentException("error parameters");
        }

        if(StringUtils.isBlank(sheetName)){
            sheetName = DEFAULT_SHEETNAME;
        }

        // 创建Workbook
        Workbook wb = createWorkbook(parseExcelType(path));

        // 创建sheet
        Sheet sheet = wb.createSheet(sheetName);

        // 冻结首行标题
        sheet.createFreezePane(0, 1, 0, 1);

        // 创建头部
        Class clazz = list.get(0).getClass();
        Field[] fields = clazz.getDeclaredFields();
        List<ExcelHeaderWrapper> headerWrappers = new ArrayList<>();
        for (Field field : fields) {
            if(!ReflectUtils.hasAnnotationPresent(field, ExcelHeader.class)){
                continue;
            }
            ExcelHeaderWrapper headerWrapper = new ExcelHeaderWrapper();
            headerWrapper.setTitle(ReflectUtils.analyzeFieldAnnotation(field, ExcelHeader.class, "title"));
            headerWrapper.setOrder(ReflectUtils.analyzeFieldAnnotation(field, ExcelHeader.class, "order"));
            headerWrapper.setWidth(ReflectUtils.analyzeFieldAnnotation(field, ExcelHeader.class, "width"));
            headerWrapper.setField(field);
            headerWrappers.add(headerWrapper);
        }
        Collections.sort(headerWrappers);   // 根据order排序
        Row headRow = sheet.createRow(0);
        CellStyle headerStyle = createCellStyle(wb, CellStylePosition.HEAD);    // 创建头部Cell样式
        for (int i = 0; i < headerWrappers.size(); i++) {
            Cell cell = headRow.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(headerWrappers.get(i).getTitle());
            sheet.setColumnWidth(i, headerWrappers.get(i).getWidth()*256);
        }

        // 创建主体
        CellStyle bodyStyle = createCellStyle(wb, CellStylePosition.BODY);    // 创建主体Cell样式
        Field field;
        out:for (int i = 0; i < list.size(); i++) {
            Row bodyRow = sheet.createRow(i+1);
            T t = list.get(i);
            in:for (int j = 0; j < headerWrappers.size(); j++) {
                field = headerWrappers.get(j).getField();
                Cell bodyCell = bodyRow.createCell(j);
                bodyCell.setCellStyle(bodyStyle);
                Object value = ReflectUtils.invokeGet(t, field.getName());
                if(value == null){
                    continue in;
                }
                String clazzName = field.getGenericType().getTypeName();
                if(GenericType.STRING.key().equals(clazzName)){
                    bodyCell.setCellValue((String) value);
                }else if(GenericType.CHAR.key().equals(clazzName)){
                    bodyCell.setCellValue(String.valueOf((char) value));
                }else if(GenericType.DATE.key().equals(clazzName)){
                    String format = com.qc.itaojin.util.DateUtil.DATE_FORMAT_1;
                    if(ReflectUtils.hasAnnotationPresent(field, DateFormat.class)){
                        format = ReflectUtils.analyzeFieldAnnotation(field, DateFormat.class, "format");
                    }
                    bodyCell.setCellValue(com.qc.itaojin.util.DateUtil.format((Date) value, format));
                }else if(GenericType.BOOLEAN.key().equals(clazzName) || GenericType.BOOLEAN_PACKAGE.key().equals(clazzName)){
                    bodyCell.setCellValue((boolean) value);
                }else if(isNumeric(clazzName)){
                    bodyCell.setCellValue(parseToDouble(clazzName, value));
                }else{
                    if(ReflectUtils.hasAnnotationPresent(field, ExcelTypeHandler.class)){
                        Class typeHandlerClass = ReflectUtils.analyzeFieldAnnotation(field, ExcelTypeHandler.class, "clazz");
                        try {
                            bodyCell.setCellValue(((com.efficient.excel.enums.ExcelTypeHandler)typeHandlerClass.newInstance()).onExport(value));
                        } catch (Exception e) {
                            throw new RuntimeException("error on parse @ExcelTypeHandler");
                        }
                    }
                }
            }
        }

        // 写数据
        writeTo(wb, path);
    }

    /**
     * 将不同类型的数字类型转换为Double
     * */
    private static double parseToDouble(String className, Object value){
        double result = -1;
        if(GenericType.BYTE.key().equals(className) || GenericType.BYTE_PACKAGE.key().equals(className)){
            result = ((Byte)value).doubleValue();
        }else if(GenericType.SHORT.key().equals(className) || GenericType.SHORT_PACKAGE.key().equals(className)){
            result = ((Short)value).doubleValue();
        }else if(GenericType.INT.key().equals(className) || GenericType.INTEGER.key().equals(className)){
            result = ((Integer)value).doubleValue();
        }else if(GenericType.LONG.key().equals(className) || GenericType.LONG_PACKAGE.key().equals(className)){
            result = ((Long)value).doubleValue();
        }else if(GenericType.DOUBLE.key().equals(className) || GenericType.DOUBLE_PACKAGE.key().equals(className)){
            result = (double) value;
        }else if(GenericType.FLOAT.key().equals(className) || GenericType.FLOAT_PACKAGE.key().equals(className)){
            result = ((Float)value).doubleValue();
        }else if(GenericType.BIGDECIMAL.key().equals(className)){
            result = ((BigDecimal)value).doubleValue();
        }

        return result;
    }

    /**
     * 判断值是否是梳子型
     * */
    private static boolean isNumeric(String className){
        if(StringUtils.isBlank(className)){
            return false;
        }

        if(GenericType.BYTE.key().equals(className) || GenericType.BYTE_PACKAGE.key().equals(className)
                || GenericType.SHORT.key().equals(className) || GenericType.SHORT_PACKAGE.key().equals(className)
                || GenericType.INT.key().equals(className) || GenericType.INTEGER.key().equals(className)
                || GenericType.LONG.key().equals(className) || GenericType.LONG_PACKAGE.key().equals(className)
                || GenericType.DOUBLE.key().equals(className) || GenericType.DOUBLE_PACKAGE.key().equals(className)
                || GenericType.FLOAT.key().equals(className) || GenericType.FLOAT_PACKAGE.key().equals(className)
                || GenericType.BIGDECIMAL.key().equals(className)){
            return true;
        }

        return false;
    }

    /**
     * 写入文件
     * */
    private static void writeTo(Workbook wb, String path){
        try(OutputStream os = new FileOutputStream(new File(path))){
            wb.write(os);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            try {
                wb.close();
            } catch (IOException e) {

            }
        }
    }

    /**
     * 创建CellStyle
     * @param wb
     * @param position 样式应用到哪个部位
     * */
    private static CellStyle createCellStyle(Workbook wb, CellStylePosition position){
        CellStyle style = wb.createCellStyle();
        // 设置单元格字体水平、垂直居中
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置单元格边框
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        // 设置单元格字体
        Font font = wb.createFont();
        font.setFontName("宋体");
        if(CellStylePosition.HEAD.equals(position)){
            font.setFontHeightInPoints((short)11);
            font.setBold(true);
        }else{
            font.setFontHeightInPoints((short)10);
        }
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
        style.setWrapText(false);
        return style;
    }

    /**
     * 判断一个double型数是否是<<真小数>>
     * 考虑到java中double类型在长度比较长时会转化为10的n次方表达式，例如 1.63543E7
     * */
    private static boolean isDecimal(double source){
        String doubleString = new Double(source).toString();
        if(!doubleString.contains("e") && !doubleString.contains("E")){
            if(doubleString.indexOf(".") != -1 && Long.parseLong(doubleString.substring(doubleString.indexOf(".")+1)) == 0){
                return false;
            }
        }else{
            String preFix = "";
            String suffix = "";
            if(doubleString.contains("e")){
                preFix = doubleString.substring(0, doubleString.indexOf("e"));
                suffix = doubleString.substring(doubleString.indexOf("e")+1);
            }else if(doubleString.contains("E")){
                preFix = doubleString.substring(0, doubleString.indexOf("E"));
                suffix = doubleString.substring(doubleString.indexOf("E")+1);
            }
            if(preFix.substring(preFix.indexOf(".")+1).length() <= Integer.parseInt(suffix)){
                return false;
            }
        }

        return true;
    }

    /**
     * 创建空的Wrokbook
     * */
    private static Workbook createWorkbook(ExcelType excelType){
        Workbook wb = null;
        if(excelType == null){
            return wb;
        }else if(ExcelType.XLS.equalsTo(excelType)){
            wb = new HSSFWorkbook();
        }else if(ExcelType.XLSX.equalsTo(excelType)){
            wb = new XSSFWorkbook();
        }

        return wb;
    }

    /**
     * 解析WorkWorkbook，区分.xls和.xlsx
     * */
    private static Workbook parseWorkbook(String path){
        try{
            if(parseExcelType(path).equalsTo(ExcelType.XLS)){
                return new HSSFWorkbook(new NPOIFSFileSystem(new File(path)));
            }
            if(parseExcelType(path).equalsTo(ExcelType.XLSX)){
                return new XSSFWorkbook(OPCPackage.open(new File(path)));
            }

            throw new IllegalArgumentException("invalid path");
        }catch (Exception e){
            e.printStackTrace();
        }

        return null;
    }

    /**
     * 根据文件路径名，或者文件名判断excel文件类型 .xls xlsx
     * */
    private static ExcelType parseExcelType(String folderName){
        if(StringUtils.isBlank(folderName)){
            return ExcelType.UNKNOWN;
        }

        if(folderName.endsWith(ExcelType.XLS.suffix())){
            return ExcelType.XLS;
        }

        if(folderName.endsWith(ExcelType.XLSX.suffix())){
            return ExcelType.XLSX;
        }

        throw new RuntimeException("invalid excel type");
    }

    /**
     * 解析Sheet
     * */
    private static Sheet parseSheet(Workbook wb, String sheetName){
        for(Sheet sheet : wb){
            if(sheetName.equals(sheet.getSheetName())){
                return sheet;
            }
        }

        throw new RuntimeException("invalid sheetName {" + sheetName + "}");
    }

    /**
     * 空sheet
     * */
    public static boolean isEmpty(Sheet sheet){
        if(sheet==null){
            return true;
        }

        if(sheet.getLastRowNum()==0 && sheet.getPhysicalNumberOfRows()==0){
            return true;
        }

        return false;
    }

}
