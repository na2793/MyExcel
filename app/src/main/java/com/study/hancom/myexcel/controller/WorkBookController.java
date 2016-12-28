package com.study.hancom.myexcel.controller;

import android.content.ContentResolver;
import android.content.Context;
import android.database.Cursor;
import android.net.Uri;
import android.provider.OpenableColumns;
import android.util.Log;

import com.study.hancom.myexcel.model.Cell;
import com.study.hancom.myexcel.model.Sheet;
import com.study.hancom.myexcel.model.WorkBook;
import com.study.hancom.myexcel.util.exception.CellOutOfSheetBoundException;
import com.study.hancom.myexcel.util.Hana;
import com.study.hancom.myexcel.util.exception.DuplicatedSheetNameException;
import com.study.hancom.myexcel.util.exception.SheetOutOfSheetListBoundsException;
import com.study.hancom.myexcel.util.listener.DataChangeListenerService;

import static com.study.hancom.myexcel.BuildConfig.DEBUG;
import static com.study.hancom.myexcel.util.Hana.HANA_FILENAME_EXTENSION;
import static com.study.hancom.myexcel.view.WorkBookView.SECTION_TYPE_RESET;

public class WorkBookController {
    private static final String TAG = "WorkBookController";

    private static final int DEFAULT_SHEET_COUNT = 3;

    private String mFileName = null;
    private WorkBook mWorkBook = null;

    public WorkBookController() {
        init();
    }

    public void init() {
        mFileName = null;
        mWorkBook = new WorkBook();
        DataChangeListenerService.pause();
        for (int i = 0; i < DEFAULT_SHEET_COUNT; i++) {
            mWorkBook.createSheet();
        }
        DataChangeListenerService.resume(SECTION_TYPE_RESET);
    }

    public String getFileName() {
        return mFileName;
    }

    public int getSheetCount() {
        return mWorkBook.getAllSheetId().length;
    }

    public Sheet addSheet() {
        Sheet sheet = mWorkBook.createSheet();

        return sheet;
    }

    public void deleteSheet(int sheetNum) {
        int[] allSheetId = mWorkBook.getAllSheetId();

        mWorkBook.deleteSheet(allSheetId[sheetNum]);
    }

    public String getSheetName(int sheetNum) {
        return getSheet(sheetNum).getName();
    }

    public void setSheetName(int sheetNum, String newSheetName) throws DuplicatedSheetNameException {
        int[] allSheetId = mWorkBook.getAllSheetId();

        mWorkBook.setSheetName(allSheetId[sheetNum], newSheetName);
    }

    public void changeSheetOrder(int sheetNum, int newSheetOrder) throws SheetOutOfSheetListBoundsException {
        int[] allSheetId = mWorkBook.getAllSheetId();

        mWorkBook.changeSheetOrder(allSheetId[sheetNum], newSheetOrder);
    }

    public Cell getCell(int sheetNum, int column, int row) throws CellOutOfSheetBoundException {
        return getSheet(sheetNum).getCell(column, row);
    }

    public Cell setCell(int sheetNum, int column, int row, String value, int type) throws CellOutOfSheetBoundException {
        Cell cell = null;

        if (value != null) {
            Sheet sheet = getSheet(sheetNum);
            if (sheet != null) {
                //**무조건 삭제
                sheet.deleteCell(column, row);
                if (value.length() > 0) {
                    cell = sheet.createCell(column, row, value, type);
                }
                if (DEBUG) {
                    Log.d(TAG, "type : " + type);
                }
            }
        }

        return cell;
    }

    // .hana 파일을 저장합니다.
    public void saveHana(Context context, String fileName) throws Exception {

        if (context == null || mWorkBook == null) {
            return;
        }

        Hana hana = new Hana();

        mFileName = hana.encodeFile(fileName, mWorkBook);
    }

    // .hana 파일을 읽습니다.
    public void readHana(Context context, Uri uri) throws Exception {
        String tempFileName = getFileName(context, uri);

        if (context == null || uri == null || !tempFileName.endsWith(HANA_FILENAME_EXTENSION)) {
            throw new Exception();
        }

        // 화면을 복구할 수 있도록 기존 workBook을 백업합니다.
        WorkBook tempWorkBook = mWorkBook.clone();

        try {
            Hana hana = new Hana();
            DataChangeListenerService.pause();
            mWorkBook = hana.decodeFile(context, uri);
            DataChangeListenerService.resume(SECTION_TYPE_RESET);
            mFileName = tempFileName;
        } catch (Exception e) {
            if (DEBUG) {
                Log.v(TAG, "restore workBook");
            }
            mWorkBook = tempWorkBook;

            throw e;
        }
    }

    private Sheet getSheet(int sheetNum) {
        int[] allSheetId = mWorkBook.getAllSheetId();

        return mWorkBook.getSheet(allSheetId[sheetNum]);
    }

    private String getFileName(Context context, Uri uri) {
        if (uri == null) {
            return null;
        }

        ContentResolver contentResolver = context.getContentResolver();
        String fileName = null;

        if (uri.getScheme().equals("content")) {
            Cursor cursor = null;
            try {
                cursor = contentResolver.query(uri, null, null, null, null);
                if (cursor != null && cursor.moveToFirst()) {
                    // 파일 이름을 받아옵니다.
                    fileName = cursor.getString(cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME));
                    if (DEBUG) {
                        Log.d(TAG, "fileName : " + fileName);
                    }
                }
            } finally {
                if (cursor != null) {
                    cursor.close();
                }
            }
        }

        return fileName;
    }
}