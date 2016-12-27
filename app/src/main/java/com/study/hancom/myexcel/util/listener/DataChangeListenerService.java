package com.study.hancom.myexcel.util.listener;

import android.util.Log;

import java.util.HashSet;

import static com.study.hancom.myexcel.BuildConfig.DEBUG;

public class DataChangeListenerService {

    private static final String TAG = "DataChangeListener";

    private static boolean mBlocked = false;

    public interface OnDataChangeListener {
        void onDataChanged(int section); // 이벤트 발생 시 처리할 작업을 정의합니다.
    }

    private static HashSet<OnDataChangeListener> dataChangeListenerList = new HashSet<>();  // 추가된 listener들을 담을 데이터 구조체 (중복 불가)

    public static void setDataChangeListener(OnDataChangeListener dataChangeListener) {
        dataChangeListenerList.add(dataChangeListener);
    }

    public static void removeDataChangeListener(OnDataChangeListener dataChangeListener) {
        dataChangeListenerList.remove(dataChangeListener);
    }

    public static void pause() {
        mBlocked = true;
    }

    public static void resume() {
        mBlocked = false;
    }

    public static void resume(int dataType) {
        mBlocked = false;
        onDataChanged(dataType);
    }

    public static void onDataChanged(int dataType) {
        if (!mBlocked) {
            if (DEBUG) {
                Log.d(TAG, "data changed");
            }
            notifyChanged(dataType);
        }
    }

    private static void notifyChanged(int dataType) {
        for (OnDataChangeListener eachDataChangeListener : dataChangeListenerList) {
            eachDataChangeListener.onDataChanged(dataType);
        }
    }
}