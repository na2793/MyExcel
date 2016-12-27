package com.study.hancom.myexcel.util.exception;

public class CellOutOfSheetBoundException extends IndexOutOfBoundsException {
    public CellOutOfSheetBoundException(String message) {
        super(message);
    }
}
