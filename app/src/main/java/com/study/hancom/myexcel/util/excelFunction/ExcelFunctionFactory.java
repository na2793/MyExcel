package com.study.hancom.myexcel.util.excelFunction;

import static com.study.hancom.myexcel.util.excelFunction.ExcelFunction.EXCEL_FUNCTION_START_CHAR;
import static com.study.hancom.myexcel.util.excelFunction.ExcelFunction.EXCEL_FUNCTION_VAR_OPENER;

public class ExcelFunctionFactory {

    public static boolean isExcelFunction(final String expression) {

        final int functionNameEndPoint = expression.indexOf(EXCEL_FUNCTION_VAR_OPENER);

        if (functionNameEndPoint > 0) {
            final String functionName = expression.substring(expression.indexOf(EXCEL_FUNCTION_START_CHAR) + 1, functionNameEndPoint);

            // 아래 구문을 추가해야 합니다.
            if (functionName.equals(SUM.EXCEL_FUNCTION_NAME_SUM)) {
                return true;
            } else if (functionName.equals(AVG.EXCEl_FUNCTION_NAME_AVG)) {
                return true;
            }
        }

        return false;
    }

    public static ExcelFunction getExcelFunction(final String expression) {

        final int functionNameEndPoint = expression.indexOf(EXCEL_FUNCTION_VAR_OPENER);

        if (functionNameEndPoint > 0) {
            final String functionName = expression.substring(expression.indexOf(EXCEL_FUNCTION_START_CHAR) + 1, functionNameEndPoint);

            // 아래 구문을 추가해야 합니다.
            if (functionName.equals(SUM.EXCEL_FUNCTION_NAME_SUM)) {
                return new SUM(expression);
            } else if (functionName.equals(AVG.EXCEl_FUNCTION_NAME_AVG)) {
                return new AVG(expression);
            }
        }

        return null;
    }
}
