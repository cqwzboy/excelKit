package com.efficient.excel.enums;

/**
 * Excel类型 旧版-.xls 新版-.xlsx
 * @author qinqin Fu
 * @since 2018-10-11
 */
public enum ExcelType {
    UNKNOWN("未知"),
    XLS(".xls"),
    XLSX(".xlsx"),
    ;

    private String suffix;

    ExcelType(String suffix){
        this.suffix = suffix;
    }

    public String suffix() {
        return this.suffix;
    }

    public boolean equalsTo(ExcelType excelType) {
        if(excelType == null){
            return false;
        }

        if(this == excelType){
            return true;
        }

        if(this.name().equals(excelType.name()) && this.suffix().equals(excelType.suffix())){
            return true;
        }

        return false;
    }
}
