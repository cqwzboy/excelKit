package com.efficient.excel.annotation;

import java.lang.annotation.*;

/**
 * Excel 枚举类类型转换器
 *
 * @author qinqin Fu
 * @since 2018-10-11
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelTypeHandler {
    Class clazz();
}
