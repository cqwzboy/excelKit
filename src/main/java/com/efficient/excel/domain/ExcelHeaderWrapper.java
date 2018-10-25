package com.efficient.excel.domain;

import lombok.Data;
import lombok.ToString;

import java.lang.reflect.Field;

/**
 * Excel头部实体类
 *
 * @author qinqin Fu
 * @since 2018-10-12
 */
@Data
@ToString
public class ExcelHeaderWrapper implements Comparable<ExcelHeaderWrapper> {
    /**
     * 标题
     * */
    private String title;
    /**
     * 标题排序，取值范围 -128到127，值越小越靠前
     * */
    private byte order;
    /**
     * 单元格宽度
     * */
    private int width;
    /**
     * 对应field对象
     * */
    private Field field;

    @Override
    public int compareTo(ExcelHeaderWrapper o) {
        return (this.getOrder() - o.getOrder())>=0 ? 1 : -1;
    }
}
