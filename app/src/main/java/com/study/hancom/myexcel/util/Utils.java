package com.study.hancom.myexcel.util;

public class Utils {
    public static boolean isTwoDimensionalArrayEmpty(Object[][] objects) {
        boolean empty = true;

        for (Object[] i : objects) {
            for (Object j : i) {
                if (j != null) {
                    empty = false;

                    break;
                }
            }
        }

        return empty;
    }

    public static boolean isNumber(String text) {
        boolean isNumber = true;

        for (int i = 0; i < text.length(); i++) {
            if (!Character.isDigit(text.charAt(i))) {
                isNumber = false;

                break;
            }
        }

        return isNumber;
    }
}
