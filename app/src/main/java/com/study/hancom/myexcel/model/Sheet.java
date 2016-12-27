package com.study.hancom.myexcel.model;

import android.util.Log;

import com.study.hancom.myexcel.util.Utils;
import com.study.hancom.myexcel.util.exception.CellOutOfSheetBoundException;
import com.study.hancom.myexcel.util.listener.DataChangeListenerService;

import java.util.HashMap;

import static com.study.hancom.myexcel.BuildConfig.DEBUG;
import static com.study.hancom.myexcel.view.WorkBookView.SECTION_TYPE_SHEET;

public class Sheet {
    private static final String TAG = "Sheet";

    private final static int SHEET_BLOCK_COLUMN = 10;
    private final static int SHEET_BLOCK_ROW = 10;

    private String mName = null;
    private int mId;

    private int mMaxColumn;
    private int mMaxRow;

    private final HashMap<Integer, Cell[][]> mCellBlockMap = new HashMap<>(); // <blockNum, cells>
    private final CellStorage mCellStorage = CellStorage.getInstance();

    // 생성자입니다. 인자의 값으로 멤버변수를 초기화해주세요.
    Sheet(String name, int maxColumn, int maxRow) {
        mName = name;
        mId = UniqueId.createUniqueId();
        mMaxColumn = maxColumn;
        mMaxRow = maxRow;
    }

    private static final class UniqueId {
        private static int uniqueId = 0;

        private static int createUniqueId() {
            return uniqueId++;
        }
    }

    // name 값을 리턴합니다.
    public String getName() {
        return mName;
    }

    // id 값을 리턴합니다.
    public int getId() {
        return mId;
    }

    public int getMaxColumn() {
        return mMaxColumn;
    }

    public int getMaxRow() {
        return mMaxRow;
    }

    // name 값을 변경합니다.
    String setName(String name) {
        String oldName = mName;

        mName = name;

        return oldName;
    }

    // cellStorage에 접근하여 cell을 얻고,
    // 이를 cellBlockMap의 적절한 위치에서 참조하도록 합니다.
    public Cell createCell(int column, int row, String value, int type) {

        if (column <= 0 || row <= 0 || column > mMaxColumn || row > mMaxRow) {
            throw new CellOutOfSheetBoundException("cell (" + column + ", " + row + ") can't be in " + mName + ". (maxColumn = " + mMaxColumn + ", maxRow = " + mMaxRow + ")");
        }

        Cell[][] cells = null;
        Cell cell = mCellStorage.createCell(mId, value, type);  // cell storage에서 인자 값을 가진 cell을 가져옵니다.
        int blockNum = getCellBlockNum(column, row);
        boolean exist = false;

        // 순환하며 sheet block의 처음 좌표 값을 키 값으로 갖는 entry를 찾습니다.
        for (int eachBlockNum : mCellBlockMap.keySet()) {
            if (eachBlockNum == blockNum) {
                exist = true;
                cells = mCellBlockMap.get(eachBlockNum);

                break;
            }
        }

        // 기존 entry를 찾지 못했다면 새로 생성합니다.
        if (!exist) {
            if (mMaxRow % SHEET_BLOCK_ROW != 0) {
                int maxBlockColumn = (int) Math.ceil((float) mMaxColumn / (float) SHEET_BLOCK_COLUMN);
                int maxBlockRow = (int) Math.ceil((float) mMaxRow / (float) SHEET_BLOCK_ROW);

                // 배열을 찾습니다.
                if (blockNum == maxBlockColumn * maxBlockRow) {
                    cells = new Cell[mMaxRow % SHEET_BLOCK_ROW][mMaxColumn % SHEET_BLOCK_COLUMN];
                } else if ((maxBlockColumn - 1) * maxBlockRow < blockNum && blockNum < maxBlockColumn * maxBlockRow) {
                    cells = new Cell[mMaxRow % SHEET_BLOCK_ROW][SHEET_BLOCK_COLUMN];
                } else if (blockNum % maxBlockColumn == 0) {
                    cells = new Cell[SHEET_BLOCK_ROW][mMaxColumn % SHEET_BLOCK_COLUMN];
                } else {
                    cells = new Cell[SHEET_BLOCK_ROW][SHEET_BLOCK_COLUMN];
                }
            } else {
                cells = new Cell[SHEET_BLOCK_ROW][SHEET_BLOCK_COLUMN];
            }

            mCellBlockMap.put(blockNum, cells);

            if (DEBUG) {
                Log.d(TAG, mName + " created new sheet block " + blockNum + ".");
                Log.d(TAG, "new sheet block size : [" + cells.length + "][" + cells[0].length + "]");
            }
        } else {
            if (DEBUG) {
                Log.d(TAG, "sheet block " + blockNum + " already existed in " + mName + ".");
            }
        }

        // cell을 배열에 넣습니다.
        int rowInBlock = (row - 1) % SHEET_BLOCK_ROW;
        int columnInBlock = (column - 1) % SHEET_BLOCK_COLUMN;
        cells[rowInBlock][columnInBlock] = cell;

        if (DEBUG) {
            Log.d(TAG, mName + " got cell " + cell + ". (" + column + ", " + row + ")");
        }

        DataChangeListenerService.onDataChanged(SECTION_TYPE_SHEET);

        // cell을 리턴합니다.
        return cell;
    }

    // cellBlockMap에 접근하여 특정 위치의 cell을 리턴합니다.
    public Cell getCell(int column, int row) {

        if (column <= 0 || row <= 0 || column > mMaxColumn || row > mMaxRow) {
            throw new CellOutOfSheetBoundException("cell (" + column + ", " + row + ") can't be in " + mName + ". (maxColumn = " + mMaxColumn + ", maxRow = " + mMaxRow + ")");
        }

        Cell cell = null;
        int rowInBlock = (row - 1) % SHEET_BLOCK_ROW;
        int columnInBlock = (column - 1) % SHEET_BLOCK_COLUMN;

        // 순환하며 sheet block의 처음 좌표 값을 키 값으로 갖는 entry를 찾습니다.
        for (int eachBlockNum : mCellBlockMap.keySet()) {
            if (eachBlockNum == getCellBlockNum(column, row)) {
                Cell[][] cells = mCellBlockMap.get(eachBlockNum);
                cell = cells[rowInBlock][columnInBlock];
            }
        }

        return cell;
    }

    // cellBlockMap이 더 이상 해당 cell을 참조하지 않도록 합니다.
    public Cell deleteCell(int column, int row) {

        if (column <= 0 || row <= 0 || column > mMaxColumn || row > mMaxRow) {
            throw new CellOutOfSheetBoundException("cell (" + column + ", " + row + ") can't be in " + mName + ". (maxColumn = " + mMaxColumn + ", maxRow = " + mMaxRow + ")");
        }

        Cell cell = null;
        int rowInBlock = (row - 1) % SHEET_BLOCK_ROW;
        int columnInBlock = (column - 1) % SHEET_BLOCK_COLUMN;

        // 순환하며 sheet block의 처음 좌표 값을 키 값으로 갖는 entry를 찾습니다.
        for (int eachBlockNum : mCellBlockMap.keySet()) {
            if (eachBlockNum == getCellBlockNum(column, row)) {
                Cell[][] cellBlock = mCellBlockMap.get(eachBlockNum);

                cell = cellBlock[rowInBlock][columnInBlock];
                cellBlock[rowInBlock][columnInBlock] = null;

                if (DEBUG) {
                    Log.d(TAG, mName + " deleted cell " + cell + ". (" + column + ", " + row + ")");
                }

                if (Utils.isTwoDimensionalArrayEmpty(cellBlock)) {
                    mCellBlockMap.remove(eachBlockNum);
                    if (DEBUG) {
                        Log.d(TAG, "sheet block " + eachBlockNum + " was empty. " + mName + " deleted sheet block.");
                    }
                }

                break;
            }
        }

        // cell storage에서도 cell을 삭제합니다.
        mCellStorage.deleteCell(mId, cell);

        DataChangeListenerService.onDataChanged(SECTION_TYPE_SHEET);

        return cell;
    }

    void clear() {
        mCellStorage.deleteAllCellsInSheet(mId);
        mCellBlockMap.clear();
    }

    // sheet block number를 리턴합니다.
    // 예시 )
    // -------------
    // | 1 | 2 | 3 |
    // | 4 | 5 | 6 |
    // | 7 | 8 | 9 |
    // -------------
    private int getCellBlockNum(int column, int row) {

        int lastBlockColumn = mMaxColumn % SHEET_BLOCK_COLUMN;
        int blockNum;

        // 배열은 (0, 0), cell은 (1, 1)을 기준으로 하므로 배열에 맞춰 감산합니다.
        --row;
        --column;

        // 쭉 펴서
        blockNum = (int) ((float) row / (float) SHEET_BLOCK_ROW) * mMaxColumn + column;

        // 마지막 column에 부족한 만큼 더하고
        blockNum += (blockNum / mMaxColumn) * (SHEET_BLOCK_COLUMN - lastBlockColumn);

        // block 단위로 갯수를 셉니다. (block number은 1부터 시작합니다.)
        blockNum = blockNum / SHEET_BLOCK_COLUMN + 1;

        return blockNum;
    }
}