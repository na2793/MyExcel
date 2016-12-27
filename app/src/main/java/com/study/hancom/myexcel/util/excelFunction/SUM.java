package com.study.hancom.myexcel.util.excelFunction;

final class SUM extends ExcelFunction {
    static final String EXCEL_FUNCTION_NAME_SUM = "SUM";
    private String mExpression = null;

    SUM(String expression) {
        mExpression = expression;
    }

    @Override
    public String getExpression() {

        return mExpression;
    }

    @Override
    public String getResult() {
        String number1 = mExpression.substring(mExpression.indexOf(EXCEL_FUNCTION_VAR_OPENER) + 1, mExpression.indexOf(EXCEL_FUNCTION_VAR_SEPARATOR));
        String number2 = mExpression.substring(mExpression.indexOf(EXCEL_FUNCTION_VAR_SEPARATOR) + 1, mExpression.length() - EXCEL_FUNCTION_VAR_CLOSER.length());

        return Double.toString((Double.parseDouble(number1) + Double.parseDouble(number2)));
    }
}
