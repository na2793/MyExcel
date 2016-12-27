package com.study.hancom.myexcel.util.excelFunction;

public abstract class ExcelFunction {
    public static final String EXCEL_FUNCTION_START_CHAR = "=";

    static final String EXCEL_FUNCTION_VAR_OPENER = "(";
    static final String EXCEL_FUNCTION_VAR_CLOSER = ")";
    static final String EXCEL_FUNCTION_VAR_SEPARATOR = ",";

    public abstract String getExpression();
    public abstract String getResult();
}
