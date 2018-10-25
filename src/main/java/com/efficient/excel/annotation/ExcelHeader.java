package com.efficient.excel.annotation;

import java.lang.annotation.*;

/**
 * Excel标题映射
 *
 * @author qinqin Fu
 * @since 2018-10-11
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelHeader {
    String title();

    byte order() default -128;

    int width() default 20;
}
