package com.study.hancom.myexcel.model;

import com.study.hancom.myexcel.util.exception.DuplicatedSheetNameException;
import com.study.hancom.myexcel.util.listener.DataChangeListenerService;

import static com.study.hancom.myexcel.view.WorkBookView.SECTION_TYPE_ALL;
import static com.study.hancom.myexcel.view.WorkBookView.SECTION_TYPE_REDRAW;
import static com.study.hancom.myexcel.view.WorkBookView.SECTION_TYPE_SHEET_MENU;

public class WorkBook implements Cloneable {
    private static final int DEFAULT_SHEET_NUM_COLUMNS = 26;   // 26 = column (A ~ Z)
    private static final int DEFAULT_SHEET_NUM_ROWS = 100;  // 100 = row (1 ~ 100)

    private final SheetManager mSheetManager = new SheetManager();

    public WorkBook() {
    }

    // sheetManager에 접근하여 인자의 값을 갖는 sheet를 생성합니다.
    public Sheet createSheet() {
        return this.createSheet(DEFAULT_SHEET_NUM_COLUMNS, DEFAULT_SHEET_NUM_ROWS);
    }

    public Sheet createSheet(int maxColumn, int maxRow) {
        Sheet sheet = mSheetManager.createSheet(maxColumn, maxRow);
        DataChangeListenerService.onDataChanged(SECTION_TYPE_REDRAW);

        return sheet;
    }

    public Sheet createSheet(String sheetName) {
        return this.createSheet(sheetName, DEFAULT_SHEET_NUM_COLUMNS, DEFAULT_SHEET_NUM_ROWS);
    }

    public Sheet createSheet(String sheetName, int maxColumn, int maxRow) {
        Sheet sheet = mSheetManager.createSheet(sheetName, maxColumn, maxRow);
        DataChangeListenerService.onDataChanged(SECTION_TYPE_REDRAW);

        return sheet;
    }

    // sheetManager에 접근하여 인자의 값을 갖는 sheet를 삭제합니다.
    public void deleteSheet(int sheetId) {
        mSheetManager.deleteSheet(sheetId);
        DataChangeListenerService.onDataChanged(SECTION_TYPE_ALL);
    }

    // sheetManager에 접근하여 생성된 sheet들을 리턴합니다.
    public int[] getAllSheetId() {
        return mSheetManager.getAllSheetId();
    }

    // sheetManager에 접근하여 인자의 값을 갖는 sheet를 리턴합니다.
    public Sheet getSheet(int sheetId) {
        return mSheetManager.getSheet(sheetId);
    }

    public void setSheetName(int sheetId, String newSheetName) throws DuplicatedSheetNameException {
        mSheetManager.setSheetName(sheetId, newSheetName);
        DataChangeListenerService.onDataChanged(SECTION_TYPE_SHEET_MENU);
    }

    public void changeSheetOrder(int sheetId, int newSheetOrder) {
        mSheetManager.changeSheetOrder(sheetId, newSheetOrder);
        DataChangeListenerService.onDataChanged(SECTION_TYPE_SHEET_MENU);
    }

    public WorkBook clone() {
        try {
            return (WorkBook) super.clone();
        } catch (CloneNotSupportedException e) {
            // this shouldn't happen, since we are Cloneable
            throw new InternalError();
        }
    }
}