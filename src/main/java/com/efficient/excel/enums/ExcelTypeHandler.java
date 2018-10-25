package com.efficient.excel.enums;

/**
 * Excel 类型转换
 *
 * @author qinqin Fu
 * @since 2018-10-11
 */
public interface ExcelTypeHandler<T> {

    T onImport(String keyword);

    String onExport(T t);

}
