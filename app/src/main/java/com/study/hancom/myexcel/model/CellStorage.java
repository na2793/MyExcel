package com.study.hancom.myexcel.model;

import android.util.Log;

import java.util.HashMap;
import java.util.concurrent.ConcurrentHashMap;

import static com.study.hancom.myexcel.BuildConfig.DEBUG;

class CellStorage {
    private static final String TAG = "CellStorage";

    private final ConcurrentHashMap<Cell, HashMap<Integer, Integer>> mCellMap = new ConcurrentHashMap<>();    // <cellInstance, <sheetId, count>>

    // 해당 객체는 여러 개 생성되어선 안 되며, 따라서 생성자에 직접 접근할 수 없어야 합니다.
    // 싱글톤으로 구현하십시오.
    private CellStorage() {}

    private static class Singleton {
        private static final CellStorage instance = new CellStorage();
    }

    // 해당 메서드를 통해 인스턴스를 가질 수 있습니다.
    static CellStorage getInstance() {
        return Singleton.instance;
    }

    // 인자의 값을 갖는 cell을 생성합니다.
    // 인자의 값을 갖는 cell이 이미 생성되어 있다면 인스턴스를 새로 생성하지 않고 기존의 것을 사용하도록 합니다.
    // Cell이 사용될 때마다 cellEntry의 해당 sheet에 대한 count 값을 늘려주십시오.
    Cell createCell(int sheetId, String cellValue, int cellType) {

        Cell cell = null;
        boolean exist = false;
        HashMap<Integer, Integer> useCountMap = null;
        int useCount = 0;

        // 순환하며 인자로 받은 value, type의 값을 키 값으로 가진 cell entry를 찾습니다.
        for (Cell eachCell : mCellMap.keySet()) {
            if (eachCell.getValue().equals(cellValue) && eachCell.getType() == cellType) {
                exist = true;
                cell = eachCell;
                useCountMap = mCellMap.get(eachCell);

                // 기존 entry의 value 값에 대하여
                // 순환하며 인자로 받은 sheetId 값을 키 값으로 가진 useCount entry를 찾습니다.
                for (Integer eachUseCount : useCountMap.keySet()) {
                    if (eachUseCount == sheetId) {
                        useCount = useCountMap.get(eachUseCount);

                        break;
                    }
                }

                break;
            }
        }

        // 기존 entry를 찾지 못했다면 새로 생성합니다.
        if (!exist) {
            cell = new Cell(cellValue, cellType);
            useCountMap = new HashMap<>();
            mCellMap.put(cell, useCountMap);

            if (DEBUG) {
                Log.d(TAG, "cell storage created new cell " + cell + ".");
            }
        } else {
            if (DEBUG) {
                Log.d(TAG, "cell " + cell + " already existed in cell storage.");
            }
        }

        // 인자로 받은 sheetId에 대한 useCount를 1회 증가시킵니다.
        useCountMap.put(sheetId, ++useCount);

        // Cell instance를 리턴합니다.
        return cell;
    }

    // 이 메서드가 호출되면 인자로 받은 해당 cell에 대한 sheet의 count 값을 줄여주십시오.
    // 모든 sheet에 대한 count 값이 0이라면 해당 cell을 더 이상 갖고 있지 않게 하십시오.
    void deleteCell(int sheetId, Cell cell) {

        if (cell != null) {
            // 순환하며 인자의 값과 동일한 키 값을 가진 entry를 찾습니다.
            for (Cell eachCell : mCellMap.keySet()) {
                if (eachCell.getValue().equals(cell.getValue()) && eachCell.getType() == cell.getType()) {
                    HashMap<Integer, Integer> useCountMap = mCellMap.get(eachCell);

                    // 기존 entry가 있다면 그에 대하여
                    // 인자로 받은 sheetId 값을 키 값으로 가진 useCount entry를 찾습니다.
                    int useCount = useCountMap.get(sheetId);

                    // useCount entry의 useCount를 감산했을 때 그 값이 1보다 작으면 entry를 삭제합니다.
                    // 아니라면 감산한 값으로 entry를 갱신합니다.
                    if (--useCount < 1) {
                        useCountMap.remove(sheetId);
                    } else {
                        useCountMap.put(sheetId, useCount);
                    }

                    // useCount를 가진 sheet가 없으면(= useCountMap이 비어있으면) entry를 삭제합니다.
                    if (useCountMap.size() < 1) {
                        mCellMap.remove(cell);

                        if (DEBUG) {
                            Log.d(TAG, "cell " + eachCell + " never used. cell storage completely deleted cell.");
                        }
                    }

                    break;
                }
            }
        }
    }

    void deleteAllCellsInSheet(int sheetId) {

        // 순환하며 인자의 값과 동일한 키 값을 가진 entry를 찾습니다.
        for (Cell eachCell : mCellMap.keySet()) {
            HashMap<Integer, Integer> useCountMap = mCellMap.get(eachCell);
            useCountMap.remove(sheetId);

            // useCount를 가진 sheet가 없으면(= useCountMap이 비어있으면) entry를 삭제합니다.
            if (useCountMap.size() < 1) {
                mCellMap.remove(eachCell);

                if (DEBUG) {
                    Log.d(TAG, "cell " + eachCell + " never used. cell storage completely deleted cell.");
                }
            }
        }
    }
}