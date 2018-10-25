package com.efficient.excel.annotation;

import java.lang.annotation.*;

/**
 * 日期序列化
 *
 * @author qinqin Fu
 * @since 2018-10-12
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface DateFormat {
    String format() default "yyyy-MM-dd";
}
