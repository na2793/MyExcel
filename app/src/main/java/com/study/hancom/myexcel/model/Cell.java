package com.study.hancom.myexcel.model;

public class Cell {

    public static final int TYPE_CELL_NUMBER = 1;
    public static final int TYPE_CELL_STRING = 2;
    public static final int TYPE_CELL_FUNCTION = 3;

    private String mValue = null;
    private int mType;

    // 생성자입니다. 인자의 값으로 멤버변수를 초기화해주세요.
    Cell(String value, int type) {
        mValue = value;
        mType = type;
    }

    // value 값을 리턴합니다.
    public String getValue() {
        return mValue;
    }

    // type 값을 리턴합니다.
    public int getType() {
        return mType;
    }

    public String toString() {
        return "[value : " + mValue + ", type : " + mType + "]";
    }
}