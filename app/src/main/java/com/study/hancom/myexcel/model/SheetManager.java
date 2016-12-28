package com.study.hancom.myexcel.model;

import com.study.hancom.myexcel.util.exception.DuplicatedSheetNameException;
import com.study.hancom.myexcel.util.exception.SheetOutOfSheetListBoundsException;

import java.util.LinkedList;

import static com.study.hancom.myexcel.model.WorkBook.DEFAULT_SHEET_MAX_COLUMN;
import static com.study.hancom.myexcel.model.WorkBook.DEFAULT_SHEET_MAX_ROW;

class SheetManager {
    private static final String DEFAULT_SHEET_NAME = "sheet";

    private final LinkedList<Sheet> mSheetList = new LinkedList<>();

    SheetManager() {}

    // 인자의 값을 갖는 sheet를 생성합니다.
    Sheet createSheet() {
        return this.createSheet(null);
    }

    Sheet createSheet(String sheetName) {

        if (sheetName == null) {
            final int sheetCount = mSheetList.size();
            int sheetNameNum = 1;
            boolean exist;

            do {
                exist = false;
                for (int i = 0; i < sheetCount; i++) {
                    Sheet tempSheet = mSheetList.get(i);
                    String tempName = DEFAULT_SHEET_NAME + sheetNameNum;

                    if (tempSheet.getName().equals(tempName)) {
                        exist = true;
                        sheetNameNum++;

                        break;
                    }
                }
            } while (exist);

            sheetName = DEFAULT_SHEET_NAME + sheetNameNum;
        }

        final Sheet sheet = new Sheet(sheetName, DEFAULT_SHEET_MAX_COLUMN, DEFAULT_SHEET_MAX_ROW);
        mSheetList.addLast(sheet);

        return sheet;
    }

    // 인자의 값을 갖는 sheet를 삭제합니다.
    Sheet deleteSheet(int sheetId) {

        Sheet sheet = null;

        for (int i = 0; i < mSheetList.size(); i++) {
            sheet = mSheetList.get(i);

            if (sheet.getId() == sheetId) {
                mSheetList.remove(i);
                sheet.clear();

                break;
            }
        }

        return sheet;
    }

    // 생성된 sheet들의 id를 리턴합니다.
    int[] getAllSheetId() {

        final int idCount = mSheetList.size();
        int[] ids = new int[idCount];

        for (int i = 0; i < idCount; i++) {
            ids[i] = mSheetList.get(i).getId();
        }

        return ids;
    }

    // 인자의 값을 갖는 sheet를 리턴합니다.
    Sheet getSheet(int sheetId) {

        Sheet sheet = null;

        for (int i = 0; i < mSheetList.size(); i++) {
            sheet = mSheetList.get(i);

            if (sheet.getId() == sheetId) {

                return sheet;
            }
        }

        return sheet;
    }

    String setSheetName(int sheetId, String newSheetName) {

        final int sheetCount = mSheetList.size();
        String oldSheetName = null;
        Sheet sheet = null;

        for (int i = 0; i < sheetCount; i++) {
            Sheet tempSheet = mSheetList.get(i);

            if (tempSheet.getName().equals(newSheetName)) {
                throw new DuplicatedSheetNameException(newSheetName + " is already existed. sheets cannot have the same name as the same workbook.");
            } else if (tempSheet.getId() == sheetId) {
                sheet = tempSheet;
            }
        }

        if (sheet != null) {
            oldSheetName = sheet.setName(newSheetName);
        }

        return oldSheetName;
    }

    // 시트를 reordering 합니다.
    void changeSheetOrder(int sheetId, int newOrder) {

        final int sheetCount = mSheetList.size();

        if (0 > newOrder || newOrder > sheetCount - 1) {
            throw new SheetOutOfSheetListBoundsException("newOrder : " + newOrder + ", listSize : " + sheetCount);
        }

        for (int i = 0; i < sheetCount; i++) {
            Sheet sheet = mSheetList.get(i);

            if (sheet.getId() == sheetId) {
                mSheetList.remove(i);
                mSheetList.add(newOrder, sheet);
            }
        }
    }
}
