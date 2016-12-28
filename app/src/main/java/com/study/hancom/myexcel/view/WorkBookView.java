package com.study.hancom.myexcel.view;

import android.content.Context;
import android.content.DialogInterface;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Paint;
import android.graphics.RectF;
import android.support.v7.app.AlertDialog;
import android.util.AttributeSet;
import android.util.Log;
import android.view.GestureDetector;
import android.view.KeyEvent;
import android.view.MotionEvent;
import android.view.ScaleGestureDetector;
import android.view.View;
import android.view.inputmethod.InputMethodManager;
import android.widget.EditText;
import android.widget.NumberPicker;
import android.widget.Toast;

import com.study.hancom.myexcel.R;
import com.study.hancom.myexcel.controller.WorkBookController;
import com.study.hancom.myexcel.model.Cell;
import com.study.hancom.myexcel.model.Sheet;
import com.study.hancom.myexcel.model.WorkBook;
import com.study.hancom.myexcel.util.Utils;
import com.study.hancom.myexcel.util.excelFunction.ExcelFunctionFactory;
import com.study.hancom.myexcel.util.exception.CellOutOfSheetBoundException;
import com.study.hancom.myexcel.util.exception.DuplicatedSheetNameException;
import com.study.hancom.myexcel.util.exception.SheetOutOfSheetListBoundsException;
import com.study.hancom.myexcel.util.listener.DataChangeListenerService;
import com.study.hancom.myexcel.util.excelFunction.ExcelFunction;

import static android.content.Context.INPUT_METHOD_SERVICE;
import static android.view.KeyEvent.KEYCODE_UNKNOWN;
import static com.study.hancom.myexcel.BuildConfig.DEBUG;
import static com.study.hancom.myexcel.model.Cell.TYPE_CELL_FUNCTION;
import static com.study.hancom.myexcel.model.Cell.TYPE_CELL_NUMBER;
import static com.study.hancom.myexcel.model.Cell.TYPE_CELL_STRING;
import static com.study.hancom.myexcel.model.WorkBook.DEFAULT_SHEET_MAX_COLUMN;
import static com.study.hancom.myexcel.model.WorkBook.DEFAULT_SHEET_MAX_ROW;
import static com.study.hancom.myexcel.util.DrawText.ALIGNMENT_CENTER;
import static com.study.hancom.myexcel.util.DrawText.ALIGNMENT_LEFT;
import static com.study.hancom.myexcel.util.DrawText.ALIGNMENT_RIGHT;
import static com.study.hancom.myexcel.util.DrawText.drawText;
import static com.study.hancom.myexcel.util.excelFunction.ExcelFunction.EXCEL_FUNCTION_START_CHAR;
import static com.study.hancom.myexcel.util.excelFunction.ExcelFunctionFactory.getExcelFunction;

/**
 * Created by Anna on 2016-11-18.
 * <p>
 * ### Hardware Acceleration Issue ###
 * 하드웨어 가속이 켜져 있을 경우,
 * invalidate 시점에서 넘긴 (갱신이 필요한) 특정 영역의 값을
 * onDraw 시점에서 제대로 받아오지 못하는 문제가 발생합니다.
 * AndroidMenifest.xml에 다음 구문을 추가하여
 * 하드웨어 가속을 끄는 것으로 해결할 수 있습니다.
 * android:hardwareAccelerated="false"
 */
public class WorkBookView extends View implements DataChangeListenerService.OnDataChangeListener {
    private static final String TAG = "WorkBookView";

    private static final int TYPE_DIALOG_SHEET_OPTION = 0;
    private static final int TYPE_DIALOG_SHEET_OPTION_CHANGE_NAME = 1;
    private static final int TYPE_DIALOG_SHEET_OPTION_CHANGE_ORDER = 2;
    private static final int TYPE_DIALOG_SHEET_OPTION_DELETE = 3;

    public static final int SECTION_TYPE_RESET = 1;
    public static final int SECTION_TYPE_REDRAW = 2;
    public static final int SECTION_TYPE_ALL = 3;
    public static final int SECTION_TYPE_SHEET = 4;
    public static final int SECTION_TYPE_SHEET_AND_HEADER = 5;
    public static final int SECTION_TYPE_SHEET_MENU = 6;

    /* COLOR */
    public static final int COLOR_WORKBOOK_BACKGROUND = Color.WHITE;

    public static final int COLOR_EDITFORM_BACKGROUND = Color.WHITE;
    public static final int COLOR_EDITFORM_TEXT = Color.BLACK;
    public static final int COLOR_EDITFORM_BORDER = Color.RED;

    public static final int COLOR_SHEET_BORDER = Color.LTGRAY;
    public static final int COLOR_CELL_TEXT = Color.BLACK;

    public static final int COLOR_SELECTION_BACKGROUND = Color.rgb(220, 250, 245);
    public static final int COLOR_SELECTION_GUIDE_BACKGROUND = Color.rgb(245, 245, 245);
    public static final int COLOR_SELECTION_START_BACKGROUND = Color.WHITE;
    public static final int COLOR_SELECTION_BORDER = Color.GRAY;

    public static final int COLOR_HEADER_BACKGROUND = Color.LTGRAY;
    public static final int COLOR_HEADER_SELECTED_BACKGROUND = Color.rgb(150, 230, 220);
    public static final int COLOR_HEADER_TEXT = Color.BLACK;

    public static final int COLOR_SHEET_MENU_BACKGROUND = Color.BLACK;
    public static final int COLOR_SHEET_MENU_BUTTON_BACKGROUND = Color.LTGRAY;
    public static final int COLOR_SHEET_MENU_TEXT = Color.BLACK;
    public static final int COLOR_SHEET_MENU_SELECTED_BACKGROUND = Color.WHITE;
    public static final int COLOR_SHEET_MENU_SELECTED_TEXT = Color.BLACK;
    public static final int COLOR_SHEET_MENU_ADD_BUTTON_BACKGROUND = Color.rgb(150, 230, 220);
    public static final int COLOR_SHEET_MENU_ADD_BUTTON_TEXT = Color.BLACK;

    private static final int SCROLL_SPEED = 1;
    private static final float DEFAULT_SCALE = 1.0f;
    private static final float DEFAULT_MIN_SCALE = 0.5f;
    private static final float DEFAULT_MAX_SCALE = 3.0f;

    private static final int DEFAULT_CELL_HEIGHT = 130;
    private static final int DEFAULT_CELL_WIDTH = 250;

    private static final int DEFAULT_SHEET_NUM = 0;

    private static final int DEFAULT_SHEET_COUNT = 3;

    /* COLOR END */

    private final Paint paint = new Paint();
    private final Paint subPaint = new Paint();
    private final Paint textPaint = new Paint();
    private final Paint subTextPaint = new Paint();

    private final RectF sheetSectionRect = new RectF();
    private final RectF sheetMenuButtonsSectionRect = new RectF();
    private final RectF sheetMenuAddButtonSectionRect = new RectF();
    private final RectF headerColumnSectionRect = new RectF();
    private final RectF headerRowSectionRect = new RectF();
    private final RectF headerCrossSectionRect = new RectF();

    private WorkBookController workBookController = null;
    private WorkBook workBook = new WorkBook();

    private int maxRow = DEFAULT_SHEET_MAX_ROW;
    private int maxColumn = DEFAULT_SHEET_MAX_COLUMN;

    private int width;
    private int height;

    private float sheetHeight;

    private float cellWidth;
    private float cellHeight;
    private float fontSize;

    private float scale = DEFAULT_SCALE; // 기본 스케일(1배)로 초기화
    private float minScale = DEFAULT_MIN_SCALE;
    private float maxScale = DEFAULT_MAX_SCALE;

    private float headerColumnHeight;
    private float headerRowWidth;

    private int currentSheetNum = DEFAULT_SHEET_NUM;

    private float sheetMenuFontSize;
    private float sheetMenuButtonHeight;
    private float sheetMenuButtonWidth;
    private float sheetMenuMargin;
    private float sheetMenuAddSheetButtonWidth;

    private boolean scaleFlag = false;
    private boolean selectionFlag = false;
    private boolean doubleTapFlag = false;
    private boolean scrollFlag = false;
    private boolean editFlag = false;

    private int selectedStartColumn = 1; // 초기 선택 영역 좌표(1, 1)로 초기화
    private int selectedStartRow = 1;
    private int selectedEndRow = selectedStartColumn;
    private int selectedEndColumn = selectedStartRow;

    private int section = SECTION_TYPE_ALL;  // 뷰 전체 영역으로 초기화

    private float moveX = 0;
    private float moveY = 0;
    private float currentSheetX = 0;
    private float currentSheetY = 0;
    private float currentSheetMenuX = 0;

    private StringBuilder inputStringBuilder;
    private final InputMethodManager inputMethodManager = (InputMethodManager) getContext().getSystemService(INPUT_METHOD_SERVICE);
    
    private final OnKeyListener keyListener = new OnKeyListener() {
        @Override
        public boolean onKey(View view, int keyCode, KeyEvent event) {
            // 버튼을 눌렀을 때
            if (event.getAction() == KeyEvent.ACTION_DOWN) {
                if (DEBUG) {
                    Log.d("text", "KeyEvent.ACTION_DOWN");
                }
            }
            // 버튼을 떼었을 때
            else if (event.getAction() == KeyEvent.ACTION_UP) {
                // 엔터 키
                switch (keyCode) {
                    case KeyEvent.KEYCODE_DEL:
                        if (DEBUG) {
                            Log.d("text", "KeyEvent.KEYCODE_DEL");
                        }
                        final int length = inputStringBuilder.length();
                        if (length > 0) {
                            inputStringBuilder.delete(length - 1, length);
                        }
                        invalidateSection(SECTION_TYPE_SHEET);

                        break;
                    case KeyEvent.KEYCODE_ENTER:
                        if (DEBUG) {
                            Log.d("text", "KeyEvent.KEYCODE_ENTER");
                        }
                        editFlag = false;
                        hideSoftKeyboard();
                        try {
                            setCell(currentSheetNum, selectedStartColumn, selectedStartRow, inputStringBuilder.toString());
                        } catch (CellOutOfSheetBoundException e) {
                            if (DEBUG) {
                                e.printStackTrace();
                            }
                            final Context context = getContext();
                            Toast.makeText(context, context.getText(R.string.message_wrong_access), Toast.LENGTH_LONG).show();  //**string.xml?
                        }

                        return true;
                    default:
                        final int unicode = event.getUnicodeChar();
                        if (unicode != KEYCODE_UNKNOWN) {
                            if (DEBUG) {
                                Log.d("text", "KeyEvent.KEYCODE_CHAR");
                            }
                            inputStringBuilder.append((char) unicode);
                            invalidateSection(SECTION_TYPE_SHEET);
                        }
                }
            }

            return false;

        }
    };

    // Zoom 이벤트 처리를 위한 제스쳐 디텍터를 정의합니다.
    private final ScaleGestureDetector scaleGestureDetector = new ScaleGestureDetector(getContext(),
            new ScaleGestureDetector.SimpleOnScaleGestureListener() {
                public boolean onScale(ScaleGestureDetector detector) {
                    if (!selectionFlag) {
                        scale *= detector.getScaleFactor();

                        // 최대/최소 수준을 지정합니다.
                        scale = Math.max(minScale, Math.min(scale, maxScale));

                        // 레이아웃을 재정의합니다.
                        requestLayout();

                        invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                    }

                    return true;
                }

                @Override
                public boolean onScaleBegin(ScaleGestureDetector detector) {

                    if (DEBUG) {
                        Log.d(TAG, "scale");
                    }

                    scaleFlag = true;

                    return super.onScaleBegin(detector);
                }

                @Override
                public void onScaleEnd(ScaleGestureDetector detector) {
                    super.onScaleEnd(detector);

                    scaleFlag = false;
                }
            });

    private final GestureDetector gestureDetector = new GestureDetector(getContext(),
            new GestureDetector.SimpleOnGestureListener() {
                public boolean onDown(MotionEvent e) {
                    // double tap 이후에 발생하는 down 이벤트는 무시합니다.
                    if (!doubleTapFlag) {
                        final float x = e.getX();
                        final float y = e.getY();

                        if (sheetSectionRect.contains(x, y)) {   // 시트 영역에 대해 처리합니다.

                            final float tempColumn = getColumn(x);
                            final float tempRow = getRow(y);

                            // 선택 영역 선택자에 포함되는 영역을 누르면(down) 플래그를 세웁니다.
                            if ((selectedStartColumn == tempColumn && selectedStartRow == tempRow ||
                                    selectedEndColumn == tempColumn && selectedEndRow == tempRow)) {
                                selectionFlag = true;
                            }
                        }
                    } else {
                        doubleTapFlag = false;
                    }

                    return true;
                }

                @Override
                public boolean onSingleTapUp(MotionEvent e) {

                    if (editFlag) {
                        editFlag = false;
                        hideSoftKeyboard();

                        try {
                            setCell(currentSheetNum, selectedStartColumn, selectedStartRow, inputStringBuilder.toString());
                        } catch (CellOutOfSheetBoundException ex) {
                            ex.printStackTrace();
                            final Context context = getContext();
                            Toast.makeText(context, context.getText(R.string.message_wrong_access), Toast.LENGTH_LONG).show();  //**string.xml?
                        }
                    }

                    final float x = e.getX();
                    final float y = e.getY();

                    if (sheetSectionRect.contains(x, y)) {   // 시트 영역에 대해 처리합니다. (영역 선택 활성화)
                        final int tempColumn = getColumn(x);
                        final int tempRow = getRow(y);

                        if (tempColumn <= maxColumn && tempRow <= maxRow) {
                            selectedStartColumn = tempColumn;
                            selectedStartRow = tempRow;
                            selectedEndColumn = selectedStartColumn;
                            selectedEndRow = selectedStartRow;

                            if (DEBUG) {
                                Log.d(TAG, "single cell selection : (" + selectedStartColumn + ", " + selectedStartRow + ")");
                            }
                            invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                        }
                    } else if (sheetMenuButtonsSectionRect.contains(x, y)) {    // 메뉴 영역에 대해 처리합니다.
                        int tempSheetNum = getSheetMenuNum(x);
                        if (tempSheetNum != currentSheetNum) {
                            if (tempSheetNum < getSheetCount()) {
                                currentSheetNum = tempSheetNum;
                                if (DEBUG) {
                                    Log.d(TAG, "currentSheetNum : " + currentSheetNum);
                                }
                                Toast.makeText(getContext(), getSheetName(currentSheetNum), Toast.LENGTH_LONG).show();
                                invalidateSection(SECTION_TYPE_REDRAW);
                            }
                        }
                    } else if (sheetMenuAddButtonSectionRect.contains(x, y)) {
                        currentSheetNum = getSheetCount();

                        // 새로운 시트가 추가된 위치로 시트 메뉴를 이동합니다.
                        if (currentSheetNum > (width / (sheetMenuButtonWidth + sheetMenuMargin)) - 1) {
                            final int changedSheetCount = getSheetCount() + 1;
                            currentSheetMenuX = width - ((changedSheetCount * sheetMenuButtonWidth) + (changedSheetCount * sheetMenuMargin) + sheetMenuAddSheetButtonWidth);    // Maximum Sheet X
                        }

                        Toast.makeText(getContext(), "시트를 추가했습니다.", Toast.LENGTH_LONG).show();  //**string.xml
                        addSheet();
                    } else if (headerCrossSectionRect.contains(x, y)) {
                        selectedStartColumn = 1;
                        selectedStartRow = 1;
                        selectedEndColumn = maxColumn;
                        selectedEndRow = maxRow;

                        invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                    } else if (headerColumnSectionRect.contains(x, y)) {
                        selectedStartColumn = getColumn(x);
                        selectedStartRow = 1;
                        selectedEndColumn = getColumn(x);
                        selectedEndRow = maxRow;

                        invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                    } else if (headerRowSectionRect.contains(x, y)) {
                        selectedStartColumn = 1;
                        selectedStartRow = getRow(y);
                        selectedEndColumn = maxColumn;
                        selectedEndRow = getRow(y);

                        invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                    }

                    return true;
                }

                public void onLongPress(MotionEvent e) {

                    final float x = e.getX();
                    final float y = e.getY();

                    if (sheetMenuButtonsSectionRect.contains(x, y)) {  // 메뉴 영역에 대해 처리합니다.
                        final int tempSheetNum = getSheetMenuNum(x);

                        if (tempSheetNum < getSheetCount()) {
                            if (currentSheetNum != tempSheetNum) {
                                currentSheetNum = tempSheetNum;
                                if (DEBUG) {
                                    Log.d(TAG, "currentSheetNum : " + currentSheetNum);
                                }

                                invalidateSection(SECTION_TYPE_REDRAW);
                            }

                            createDialog(TYPE_DIALOG_SHEET_OPTION).show();
                        }
                    }
                }

                public boolean onDoubleTap(MotionEvent e) {

                    final float x = e.getX();
                    final float y = e.getY();

                    doubleTapFlag = true;

                    if (sheetSectionRect.contains(x, y)) {   // 시트 영역에 대해 처리합니다. (편집 활성화)
                        final int tempColumn = getColumn(x);
                        final int tempRow = getRow(y);

                        if (tempColumn == selectedStartColumn && tempRow == selectedStartRow) {
                            selectedEndColumn = selectedStartColumn;
                            selectedEndRow = selectedStartRow;

                            try {
                                final Cell cell = getCell(currentSheetNum, selectedStartColumn, selectedStartRow);

                                if (cell != null) {
                                    inputStringBuilder = new StringBuilder(cell.getValue());
                                } else {
                                    inputStringBuilder = new StringBuilder();
                                }

                                editFlag = true;
                                selectionFlag = false;

                                if (DEBUG) {
                                    Log.d(TAG, "edit start : (" + selectedStartColumn + ", " + selectedStartRow + ")");
                                }

                                showSoftKeyboard();
                                invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                            } catch (CellOutOfSheetBoundException ex) {
                                ex.printStackTrace();
                                final Context context = getContext();
                                Toast.makeText(context, context.getText(R.string.message_wrong_access), Toast.LENGTH_LONG).show();  //**string.xml?
                            }
                        }
                    }

                    return true;
                }

                public boolean onScroll(MotionEvent e1, MotionEvent e2, float distanceX, float distanceY) {

                    if (!scaleFlag) {
                        final float x1 = e1.getX();
                        final float y1 = e1.getY();
                        final float x2 = e2.getX();
                        final float y2 = e2.getY();

                        scrollFlag = true;

                        if (!editFlag && selectionFlag && sheetSectionRect.contains(x1, y1)) { // 선택 영역에 대해 처리합니다.
                            final int tempColumn = selectedEndColumn;
                            final int tempRow = selectedEndRow;
                            selectedEndColumn = getColumn(x2);
                            selectedEndRow = getRow(y2);
                            if ((tempColumn != selectedEndColumn || tempRow != selectedEndRow)) {
                                invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                            }
                        } else if (!editFlag && headerColumnSectionRect.contains(x1, y1)) {
                            selectedEndColumn = getColumn(x2);
                            selectedEndRow = maxRow;

                            invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                        } else if (!editFlag && headerRowSectionRect.contains(x1, y1)) {
                            selectedEndColumn = maxColumn;
                            selectedEndRow = getRow(y2);

                            invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                        } else {
                            moveX = distanceX * -1;
                            moveY = distanceY * -1;
                            if (sheetSectionRect.contains(x1, y1)) {          // 시트 영역에 대해 처리합니다.
                                moveSectionView(SECTION_TYPE_SHEET_AND_HEADER);
                                invalidateSection(SECTION_TYPE_SHEET_AND_HEADER);
                            } else if (sheetMenuButtonsSectionRect.contains(x1, y1)) {  // 메뉴 영역에 대해 처리합니다.
                                moveSectionView(SECTION_TYPE_SHEET_MENU);
                                invalidateSection(SECTION_TYPE_SHEET_MENU);
                            }
                        }
                    }

                    return true;
                }
            });

    // 세 번째 생성자가 실행될 때까지 인자를 넘깁니다.
    public WorkBookView(Context context) {
        this(context, null);
    }

    public WorkBookView(Context context, AttributeSet attrs) {
        this(context, attrs, 0);
    }

    public WorkBookView(Context context, AttributeSet attrs, int defStyleAttr) {
        super(context, attrs, defStyleAttr);

        init();
        setOnKeyListener(keyListener);
    }

    private void init() {
        workBook = new WorkBook();

        for (int i = 0; i < DEFAULT_SHEET_COUNT; i++) {
            workBook.createSheet();
        }
    }
    
    @Override
    protected void onMeasure(int widthMeasureSpec, int heightMeasureSpec) {

        // 실제 구현상의 height를 구합니다.
        switch (MeasureSpec.getMode(heightMeasureSpec)) {
            case MeasureSpec.UNSPECIFIED: // mode가 셋팅되지 않은 크기가 넘어올 때
                height = heightMeasureSpec;

                break;
            case MeasureSpec.AT_MOST: // wrap_content (뷰 내부의 크기에 따라 크기가 달라짐)
                // set default (do nothing)

                break;
            case MeasureSpec.EXACTLY: // fill_parent, match_parent (외부에서 이미 크기가 지정되었음)
                height = MeasureSpec.getSize(heightMeasureSpec);

                break;
        }

        // 실제 구현상의 width를 구합니다.
        switch (MeasureSpec.getMode(widthMeasureSpec)) {
            case MeasureSpec.UNSPECIFIED:
                width = widthMeasureSpec;

                break;
            case MeasureSpec.AT_MOST: // wrap_content
                // set default (do nothing)

                break;
            case MeasureSpec.EXACTLY: // fill_parent, match_parent
                width = MeasureSpec.getSize(widthMeasureSpec);

                break;
        }

        setMeasuredDimension(width, height);
        setViewAttributeValue();  // 속성값들을 초기화합니다.
    }

    @Override
    protected void onLayout(boolean changed, int left, int top, int right, int bottom) {
        super.onLayout(changed, left, top, right, bottom);

        sheetSectionRect.set(headerRowWidth * scale, headerColumnHeight * scale, width, sheetHeight);
        sheetMenuButtonsSectionRect.set(0, sheetHeight, width - sheetMenuAddSheetButtonWidth, height);
        sheetMenuAddButtonSectionRect.set(width - sheetMenuAddSheetButtonWidth, height - sheetMenuButtonHeight, width, height);
        headerColumnSectionRect.set(headerRowWidth * scale, 0, width, headerColumnHeight * scale);
        headerRowSectionRect.set(0, headerColumnHeight * scale, headerRowWidth * scale, sheetHeight);
        headerCrossSectionRect.set(0, 0, headerRowWidth * scale, headerColumnHeight * scale);
    }

    @Override
    protected void onFinishInflate() {
        super.onFinishInflate();
    }

    public void setController(WorkBookController workBookController) {
        this.workBookController = workBookController;
        this.workBook = null;
    }

    @Override
    public void onDataChanged(int sectionType) {
        invalidateSection(sectionType);
    }

    @Override
    protected void onDraw(Canvas canvas) {

        // 레이아웃 전체의 바탕색을 흰색으로 처리합니다.
        canvas.drawColor(COLOR_WORKBOOK_BACKGROUND);

        if (maxColumn == 0 || maxRow == 0) {
            return;
        }

        //**--특정 영역만 갱신합니다.-- invalidate 문제로 전체 영역으로 고정합니다.
        section = SECTION_TYPE_ALL;

        // 시트(셀 모음) 출력부
        if (section == SECTION_TYPE_ALL || section == SECTION_TYPE_SHEET_AND_HEADER || section == SECTION_TYPE_SHEET) {
            drawSelection(canvas);
            drawSheetLine(canvas);
            drawCellText(canvas);

            if (editFlag) {
                drawEditForm(canvas);
            }
        }

        // 시트 헤더 출력부
        if (section == SECTION_TYPE_ALL || section == SECTION_TYPE_SHEET_AND_HEADER) {
            drawSheetHeaderColumn(canvas);
            drawSheetHeaderRow(canvas);
            drawSheetHeaderCross(canvas);
        }

        // 시트 메뉴 출력부
        if (section == SECTION_TYPE_ALL || section == SECTION_TYPE_SHEET_MENU) {
            drawSheetMenuButtons(canvas);
        }

        if (section == SECTION_TYPE_ALL) {
            drawAddSheetMenuButton(canvas);
        }
    }

    private void drawEditForm(Canvas canvas) {

        final String inputString = inputStringBuilder.toString();

        canvas.save();

        canvas.clipRect(sheetSectionRect);
        canvas.scale(scale, scale);
        canvas.translate(currentSheetX + headerRowWidth, currentSheetY + headerColumnHeight);

        paint.setStyle(Paint.Style.FILL);
        paint.setColor(COLOR_EDITFORM_BACKGROUND);

        textPaint.setColor(COLOR_EDITFORM_TEXT);
        textPaint.setTextSize(fontSize);

        final float textWidth = textPaint.measureText(inputString);
        final int grabbedCellCount = (int) (textWidth / cellWidth);

        canvas.drawRect((selectedStartColumn - 1) * cellWidth, (selectedStartRow - 1) * cellHeight,
                (selectedStartColumn + grabbedCellCount) * cellWidth, selectedStartRow * cellHeight, paint);

        paint.setStyle(Paint.Style.STROKE);
        paint.setStrokeWidth(5);
        paint.setColor(COLOR_EDITFORM_BORDER);

        canvas.drawRect((selectedStartColumn - 1) * cellWidth, (selectedStartRow - 1) * cellHeight,
                (selectedStartColumn + grabbedCellCount) * cellWidth, selectedStartRow * cellHeight, paint);

        drawText(inputString, ALIGNMENT_LEFT, (selectedStartColumn - 1) * cellWidth, (selectedStartRow - 1) * cellHeight, cellWidth, cellHeight, textPaint, canvas, false);

        paint.reset();
        textPaint.reset();
        canvas.restore();
    }

    private void drawSheetLine(Canvas canvas) {

        canvas.save();

        canvas.clipRect(sheetSectionRect);
        canvas.scale(scale, scale);
        canvas.translate(currentSheetX + headerRowWidth, currentSheetY + headerColumnHeight);

        paint.setStyle(Paint.Style.STROKE);
        paint.setColor(COLOR_SHEET_BORDER);

        for (int i = 1; i < maxColumn; i++) {
            canvas.drawLine(i * cellWidth, 0, i * cellWidth, maxRow * cellHeight, paint);
        }

        for (int i = 1; i < maxRow; i++) {
            canvas.drawLine(0, i * cellHeight, maxColumn * cellWidth, i * cellHeight, paint);
        }

        paint.reset();
        canvas.restore();
    }

    private void drawCellText(Canvas canvas) {

        canvas.save();

        canvas.clipRect(sheetSectionRect);
        canvas.scale(scale, scale);
        canvas.translate(currentSheetX + headerRowWidth, currentSheetY + headerColumnHeight);

        paint.setColor(COLOR_CELL_TEXT);
        paint.setTextSize(fontSize);

        try {
            for (int row = 1; row < maxRow; row++) {
                for (int column = 1; column < maxColumn; column++) {
                    Cell cell = getCell(currentSheetNum, column, row);

                    if (cell != null) {
                        final int cellType = cell.getType();
                        String cellValue = cell.getValue();
                        int alignment;

                        switch (cellType) {
                            case TYPE_CELL_NUMBER:
                                // no break ; TYPE_CELL_FUNCTION
                            case TYPE_CELL_FUNCTION:
                                alignment = ALIGNMENT_RIGHT;
                                ExcelFunction excelFunction = getExcelFunction(cellValue);
                                if (excelFunction != null) {
                                    cellValue = excelFunction.getResult();
                                }

                                break;
                            case TYPE_CELL_STRING:
                                // no break ; default
                            default:
                                alignment = ALIGNMENT_LEFT;

                                break;
                        }

                        drawText(cellValue, alignment, (column - 1) * cellWidth, (row - 1) * cellHeight, cellWidth, cellHeight, paint, canvas, true);
                    }
                }
            }
        } catch (CellOutOfSheetBoundException e) {

            if (DEBUG) {
                e.printStackTrace();
            }

            final Context context = getContext();
            Toast.makeText(context, context.getText(R.string.message_wrong_access), Toast.LENGTH_LONG).show();  //**string.xml?
        } finally {
            paint.reset();
            canvas.restore();
        }
    }

    private void drawSelection(Canvas canvas) {

        canvas.save();

        canvas.clipRect(sheetSectionRect);
        canvas.scale(scale, scale);
        canvas.translate(currentSheetX + headerRowWidth, currentSheetY + headerColumnHeight);

        int startColumn;
        int endColumn;
        int startRow;
        int endRow;

        if (selectedStartColumn < selectedEndColumn) {
            startColumn = selectedStartColumn;
            endColumn = selectedEndColumn;
        } else {
            startColumn = selectedEndColumn;
            endColumn = selectedStartColumn;
        }

        if (selectedStartRow < selectedEndRow) {
            startRow = selectedStartRow;
            endRow = selectedEndRow;
        } else {
            startRow = selectedEndRow;
            endRow = selectedStartRow;
        }

        // 최대 범위 지정
        if (startColumn < 1) {
            startColumn = 1;
        }
        if (startRow < 1) {
            startRow = 1;
        }

        if (endColumn >= maxColumn) {
            endColumn = maxColumn;
        }
        if (endRow >= maxRow) {
            endRow = maxRow;
        }

        // 선택 영역에 대한 가이드 라인을 그립니다.
        paint.setStyle(Paint.Style.FILL);
        paint.setColor(COLOR_SELECTION_GUIDE_BACKGROUND);

        canvas.drawRect(0, (startRow - 1) * cellHeight,
                maxColumn * cellWidth, endRow * cellHeight, paint);
        canvas.drawRect((startColumn - 1) * cellWidth, 0,
                endColumn * cellWidth, maxRow * cellHeight, paint);

        // 실제 선택 영역을 그립니다.
        paint.setStyle(Paint.Style.FILL);
        paint.setColor(COLOR_SELECTION_BACKGROUND);

        canvas.drawRect((startColumn - 1) * cellWidth, (startRow - 1) * cellHeight,
                endColumn * cellWidth, endRow * cellHeight, paint);

        // 선택 영역의 시작 셀은 흰 바탕으로 표시합니다.
        if (startRow != endRow || startColumn != endColumn) {
            paint.setStyle(Paint.Style.FILL);
            paint.setColor(COLOR_SELECTION_START_BACKGROUND);

            canvas.drawRect((selectedStartColumn - 1) * cellWidth, (selectedStartRow - 1) * cellHeight,
                    selectedStartColumn * cellWidth, selectedStartRow * cellHeight, paint);
        }

        paint.setStyle(Paint.Style.STROKE);
        paint.setStrokeWidth(5);
        paint.setColor(COLOR_SELECTION_BORDER);

        canvas.drawRect((startColumn - 1) * cellWidth, (startRow - 1) * cellHeight,
                endColumn * cellWidth, endRow * cellHeight, paint);

        paint.reset();
        canvas.restore();
    }

    private void drawSheetHeaderColumn(Canvas canvas) {

        canvas.save();

        canvas.clipRect(headerColumnSectionRect);
        canvas.scale(scale, scale);
        canvas.translate(currentSheetX + headerRowWidth, 0);

        canvas.drawColor(COLOR_HEADER_BACKGROUND);

        paint.setStyle(Paint.Style.FILL);
        paint.setColor(COLOR_HEADER_SELECTED_BACKGROUND);

        textPaint.setTextSize(fontSize);
        textPaint.setColor(COLOR_HEADER_TEXT);

        for (int i = 0; i < maxColumn; i++) {
            if (selectedStartColumn <= i + 1 && i + 1 <= selectedEndColumn || selectedEndColumn <= i + 1 && i + 1 <= selectedStartColumn) {
                canvas.drawRect(i * cellWidth, 0, (i + 1) * cellWidth, headerColumnHeight, paint);
                drawText(Integer.toString(i + 1), ALIGNMENT_CENTER, i * cellWidth, 0, cellWidth, headerColumnHeight, textPaint, canvas, true);
            } else {
                drawText(Integer.toString(i + 1), ALIGNMENT_CENTER, i * cellWidth, 0, cellWidth, headerColumnHeight, textPaint, canvas, true);
            }
        }

        paint.reset();
        textPaint.reset();
        canvas.restore();
    }

    private void drawSheetHeaderRow(Canvas canvas) {

        canvas.save();

        canvas.clipRect(headerRowSectionRect);
        canvas.scale(scale, scale);
        canvas.translate(0, currentSheetY + headerColumnHeight);

        canvas.drawColor(COLOR_HEADER_BACKGROUND);

        paint.setStyle(Paint.Style.FILL);
        paint.setColor(COLOR_HEADER_SELECTED_BACKGROUND);

        textPaint.setTextSize(fontSize);
        textPaint.setColor(COLOR_HEADER_TEXT);

        for (int i = 0; i < maxRow; i++) {
            if (selectedStartRow <= i + 1 && i + 1 <= selectedEndRow || selectedEndRow <= i + 1 && i + 1 <= selectedStartRow) {
                canvas.drawRect(0, i * cellHeight, headerRowWidth, (i + 1) * cellHeight, paint);
                drawText(Integer.toString(i + 1), ALIGNMENT_CENTER, 0, i * cellHeight, headerRowWidth, cellHeight, textPaint, canvas, true);
            } else {
                drawText(Integer.toString(i + 1), ALIGNMENT_CENTER, 0, i * cellHeight, headerRowWidth, cellHeight, textPaint, canvas, true);
            }
        }

        paint.reset();
        textPaint.reset();
        canvas.restore();
    }

    private void drawSheetHeaderCross(Canvas canvas) {

        canvas.save();

        canvas.clipRect(headerCrossSectionRect);
        canvas.scale(scale, scale);
        canvas.drawColor(COLOR_HEADER_BACKGROUND);

        canvas.restore();
    }

    private void drawSheetMenuButtons(Canvas canvas) {

        canvas.save();

        canvas.clipRect(sheetMenuButtonsSectionRect);
        canvas.translate(currentSheetMenuX, 0);

        canvas.drawColor(COLOR_SHEET_MENU_BACKGROUND);

        paint.setStyle(Paint.Style.FILL);
        paint.setColor(COLOR_SHEET_MENU_BUTTON_BACKGROUND);

        textPaint.setTextSize(sheetMenuFontSize);
        textPaint.setColor(COLOR_SHEET_MENU_TEXT);

        subPaint.setStyle(Paint.Style.FILL);
        subPaint.setColor(COLOR_SHEET_MENU_SELECTED_BACKGROUND);

        subTextPaint.setTextSize(sheetMenuFontSize);
        subTextPaint.setFakeBoldText(true);
        subTextPaint.setColor(COLOR_SHEET_MENU_SELECTED_TEXT);

        String[] allSheetName = new String[getSheetCount()];

        for (int i = 0; i < allSheetName.length; i++) {
            allSheetName[i] = getSheetName(i);
        }

        for (int i = 0; i < allSheetName.length; i++) {
            if (i == currentSheetNum) {
                canvas.drawRect((i * sheetMenuButtonWidth) + (i * sheetMenuMargin), height - sheetMenuButtonHeight, ((i + 1) * sheetMenuButtonWidth) + (i * sheetMenuMargin), height, subPaint);
                drawText(allSheetName[i], ALIGNMENT_CENTER, i * sheetMenuButtonWidth + (sheetMenuMargin * i), height - sheetMenuButtonHeight, sheetMenuButtonWidth, sheetMenuButtonHeight, subTextPaint, canvas, true);
            } else {
                canvas.drawRect((i * sheetMenuButtonWidth) + (i * sheetMenuMargin), height - sheetMenuButtonHeight, ((i + 1) * sheetMenuButtonWidth) + (i * sheetMenuMargin), height, paint);
                drawText(allSheetName[i], ALIGNMENT_CENTER, i * sheetMenuButtonWidth + (sheetMenuMargin * i), height - sheetMenuButtonHeight, sheetMenuButtonWidth, sheetMenuButtonHeight, subTextPaint, canvas, true);
            }
        }

        paint.reset();
        textPaint.reset();
        subPaint.reset();
        subTextPaint.reset();

        canvas.restore();
    }

    private void drawAddSheetMenuButton(Canvas canvas) {

        canvas.save();

        canvas.clipRect(sheetMenuAddButtonSectionRect);
        canvas.translate(0, 0);

        // 객체를 재활용하기 전에 우선 리셋해줍니다.
        paint.reset();          // 기본 시트 메뉴
        textPaint.reset();      // 시트 메뉴 텍스트

        paint.setStyle(Paint.Style.FILL);
        paint.setColor(COLOR_SHEET_MENU_ADD_BUTTON_BACKGROUND);

        textPaint.setTextSize(sheetMenuFontSize);
        textPaint.setColor(COLOR_SHEET_MENU_ADD_BUTTON_TEXT);

        canvas.drawRect(width - sheetMenuAddSheetButtonWidth, height - sheetMenuButtonHeight, width, height, paint);
        drawText(getContext().getString(R.string.button_add_sheet), ALIGNMENT_CENTER, width - sheetMenuAddSheetButtonWidth, height - sheetMenuButtonHeight, sheetMenuAddSheetButtonWidth, sheetMenuButtonHeight, textPaint, canvas, false);

        paint.reset();
        textPaint.reset();

        canvas.restore();
    }    

    private void setViewAttributeValue() {
        
        cellHeight = DEFAULT_CELL_HEIGHT;
        cellWidth = DEFAULT_CELL_WIDTH;

        headerColumnHeight = cellHeight;
        headerRowWidth = cellHeight;

        sheetMenuButtonHeight = cellHeight * 3 / 2;
        sheetMenuButtonWidth = cellWidth * 3 / 2;
        sheetMenuMargin = cellWidth / 10;
        sheetMenuAddSheetButtonWidth = sheetMenuButtonWidth / 2;

        sheetHeight = height - sheetMenuButtonHeight;

        fontSize = cellHeight * 2 / 3;
        sheetMenuFontSize = fontSize;
    }

    private void invalidateSection(int sectionType) {

        switch (sectionType) {
            case SECTION_TYPE_RESET:
                if (DEBUG) {
                    Log.d(TAG, "invalidate reset");
                }
                currentSheetNum = 0;
                currentSheetMenuX = 0;
                // no break

            case SECTION_TYPE_REDRAW:
                if (DEBUG) {
                    Log.d(TAG, "invalidate redraw");
                }
                currentSheetX = 0;
                currentSheetY = 0;
                selectedStartColumn = 1;
                selectedStartRow = 1;
                selectedEndColumn = selectedStartColumn;
                selectedEndRow = selectedStartRow;
                editFlag = false;
                // no break

            case SECTION_TYPE_ALL:
                if (DEBUG) {
                    Log.d(TAG, "invalidate all");
                }
                invalidate();

                break;
            case SECTION_TYPE_SHEET:
                if (DEBUG) {
                    Log.d(TAG, "invalidate sheet : " + sheetSectionRect);
                }
                invalidate((int) sheetSectionRect.left, (int) sheetSectionRect.top, (int) sheetSectionRect.right, (int) sheetSectionRect.bottom);

                break;
            case SECTION_TYPE_SHEET_MENU:
                if (DEBUG) {
                    Log.d(TAG, "invalidate menu buttons : " + sheetMenuButtonsSectionRect);
                }
                invalidate((int) sheetMenuButtonsSectionRect.left, (int) sheetMenuButtonsSectionRect.top, (int) sheetMenuButtonsSectionRect.right, (int) sheetMenuButtonsSectionRect.bottom);

                break;
            case SECTION_TYPE_SHEET_AND_HEADER:
                if (DEBUG) {
                    Log.d(TAG, "invalidate sheet and header : " + headerCrossSectionRect);
                }
                invalidate((int) headerCrossSectionRect.left, (int) headerCrossSectionRect.top, (int) sheetSectionRect.right, (int) sheetSectionRect.bottom);

                break;
        }

        this.section = sectionType;
    }

    private void moveSectionView(int sectionType) {
        switch (sectionType) {
            case SECTION_TYPE_SHEET_AND_HEADER:
                // 스크롤 이벤트를 통해 얻은 이동값으로 시트 좌표의 기준점을 구합니다.
                currentSheetX += moveX / SCROLL_SPEED;
                currentSheetY += moveY / SCROLL_SPEED;

                // 시트가 그려진 범위 외는 출력하지 않도록 제한합니다.
                if (currentSheetX > 0) {
                    currentSheetX = 0;
                }
                if (currentSheetY > 0) {
                    currentSheetY = 0;
                }

                float tempSheetX = sheetSectionRect.width() / scale - (cellWidth * maxColumn);
                float tempSheetY = sheetSectionRect.height() / scale - (cellHeight * maxRow);

                if (currentSheetX < tempSheetX) {
                    currentSheetX = tempSheetX;
                }
                if (currentSheetY < tempSheetY) {
                    currentSheetY = tempSheetY;
                }

                break;
            case SECTION_TYPE_SHEET_MENU:
                int sheetCount = getSheetCount();
                float maxSheetMenuX = width - ((sheetCount * sheetMenuButtonWidth) + (sheetCount * sheetMenuMargin) + sheetMenuAddSheetButtonWidth);
                // 스크롤 이벤트를 통해 얻은 이동값으로 시트 메뉴 좌표의 기준점을 구합니다.
                currentSheetMenuX += moveX / SCROLL_SPEED;

                // 시트 메뉴가 그려진 범위 외는 출력하지 않도록 제한합니다.
                if (currentSheetMenuX > 0) {
                    currentSheetMenuX = 0;
                }
                if (width - sheetMenuAddSheetButtonWidth > (sheetCount * sheetMenuButtonWidth) + (sheetCount * sheetMenuMargin)) {
                    currentSheetMenuX = 0;
                } else if (currentSheetMenuX < maxSheetMenuX) {
                    currentSheetMenuX = maxSheetMenuX;
                }

                break;
        }
    }

    // 뷰포트의 절대좌표 X를 가지고 sheet menu num을 구합니다.
    private int getSheetMenuNum(float x) {
        return (int) (((x + (sheetMenuButtonWidth + sheetMenuMargin) - currentSheetMenuX)) / (sheetMenuButtonWidth + sheetMenuMargin) - 1);
    }

    // 뷰포트의 절대좌표 X를 가지고 column을 구합니다.
    private int getColumn(float x) {
        return (int) (((x + (headerColumnHeight * scale) - (currentSheetX * scale)) / (cellWidth * scale) - (int) ((headerRowWidth * scale) / (cellWidth * scale))));
    }

    // 뷰포트의 절대좌표 Y를 가지고 row를 구합니다.
    private int getRow(float y) {
        return (int) (((y + (headerRowWidth * scale) - (currentSheetY * scale)) / (cellHeight * scale) - (int) ((headerRowWidth * scale) / (cellHeight * scale))));
    }

    private AlertDialog createDialog(int type) {

        final String currentSheetName = getSheetName(currentSheetNum);

        switch (type) {
            case TYPE_DIALOG_SHEET_OPTION:
                final String[] sheetOption = {getContext().getString(R.string.button_sheet_change_name),
                        getContext().getString(R.string.button_sheet_change_order), getContext().getString(R.string.button_sheet_delete)};

                return new AlertDialog.Builder(getContext())
                        .setTitle(currentSheetName)
                        .setNegativeButton(getContext().getString(R.string.button_cancel), new DialogInterface.OnClickListener() {
                            @Override
                            public void onClick(DialogInterface dialog, int which) {
                                dialog.dismiss();     // 닫기
                            }
                        })
                        .setItems(sheetOption, // 리스트 목록에 사용할 배열
                                new DialogInterface.OnClickListener() {
                                    @Override
                                    public void onClick(DialogInterface dialog, int which) {
                                        switch (which) {
                                            case 0:
                                                createDialog(TYPE_DIALOG_SHEET_OPTION_CHANGE_NAME).show();
                                                break;
                                            case 1:
                                                createDialog(TYPE_DIALOG_SHEET_OPTION_CHANGE_ORDER).show();
                                                break;
                                            case 2:
                                                createDialog(TYPE_DIALOG_SHEET_OPTION_DELETE).show();
                                                break;
                                        }
                                    }
                                }).create();
            case TYPE_DIALOG_SHEET_OPTION_CHANGE_NAME:
                final EditText editSheetName = new EditText(getContext());
                editSheetName.setText(currentSheetName);
                editSheetName.setSelection(currentSheetName.length());
                return new AlertDialog.Builder(getContext())
                        .setTitle(currentSheetName)
                        .setMessage(getContext().getString(R.string.message_sheet_change_name) + " :")
                        .setView(editSheetName)
                        .setPositiveButton(getContext().getString(R.string.button_ok), new DialogInterface.OnClickListener() {
                            @Override
                            public void onClick(DialogInterface dialog, int which) {
                                final Context context = getContext();
                                String inputSheetName = editSheetName.getText().toString();
                                if (inputSheetName.length() > 0) {
                                    if (!inputSheetName.equals(currentSheetName)) {
                                        try {
                                            setSheetName(currentSheetNum, inputSheetName);
                                            Toast.makeText(context, context.getText(R.string.message_sheet_change_name_success), Toast.LENGTH_LONG).show();  //**string.xml?
                                        } catch (DuplicatedSheetNameException e) {
                                            if (DEBUG) {
                                                e.printStackTrace();
                                            }
                                            Toast.makeText(context, context.getText(R.string.message_sheet_change_name_duplicated_name), Toast.LENGTH_LONG).show();  //**string.xml?
                                        } finally {
                                            dialog.dismiss();     // 닫기
                                        }
                                    } else {
                                        Toast.makeText(context, context.getText(R.string.message_sheet_change_name_same_name), Toast.LENGTH_LONG).show();  //**string.xml?
                                    }
                                } else {
                                    Toast.makeText(context, context.getText(R.string.message_sheet_change_name_too_short), Toast.LENGTH_LONG).show();  //**string.xml?
                                }
                            }
                        })
                        .setNegativeButton(getContext().getString(R.string.button_cancel), new DialogInterface.OnClickListener() {
                            @Override
                            public void onClick(DialogInterface dialog, int which) {
                                dialog.dismiss();     // 닫기
                            }
                        })
                        .create();
            case TYPE_DIALOG_SHEET_OPTION_CHANGE_ORDER:
                final NumberPicker editSheetOrder = new NumberPicker(getContext());
                editSheetOrder.setMinValue(1);
                editSheetOrder.setMaxValue(getSheetCount());
                editSheetOrder.setValue(currentSheetNum + 1);
                return new AlertDialog.Builder(getContext())
                        .setTitle(currentSheetName)
                        .setMessage(getContext().getString(R.string.message_sheet_change_order)
                                + " (" + getContext().getString(R.string.message_sheet_change_order_current)
                                + " " + (currentSheetNum + 1) + "/" + getSheetCount() + ") :")
                        .setView(editSheetOrder)
                        .setPositiveButton(getContext().getString(R.string.button_ok), new DialogInterface.OnClickListener() {
                            @Override
                            public void onClick(DialogInterface dialog, int which) {
                                final Context context = getContext();
                                int inputSheetOrder = editSheetOrder.getValue() - 1;
                                if (currentSheetNum != inputSheetOrder) {
                                    try {
                                        changeSheetOrder(currentSheetNum, inputSheetOrder);
                                        currentSheetNum = inputSheetOrder;
                                        Toast.makeText(context, context.getText(R.string.message_sheet_change_order_success), Toast.LENGTH_LONG).show();  //**string.xml?
                                    } catch (SheetOutOfSheetListBoundsException e) {
                                        if (DEBUG) {
                                            e.printStackTrace();
                                        }
                                        Toast.makeText(context, context.getText(R.string.message_wrong_access), Toast.LENGTH_LONG).show();  //**string.xml?
                                    }
                                } else {
                                    Toast.makeText(context, context.getText(R.string.message_sheet_change_order_same_order), Toast.LENGTH_LONG).show();  //**string.xml?
                                }
                                dialog.dismiss();
                            }
                        })
                        .setNegativeButton(getContext().getString(R.string.button_cancel), new DialogInterface.OnClickListener() {
                            @Override
                            public void onClick(DialogInterface dialog, int which) {
                                dialog.dismiss();
                            }
                        })
                        .create();
            case TYPE_DIALOG_SHEET_OPTION_DELETE:
                return new AlertDialog.Builder(getContext())
                        .setTitle(currentSheetName)
                        .setMessage(getContext().getString(R.string.message_sheet_delete))
                        .setPositiveButton(getContext().getString(R.string.button_yes), new DialogInterface.OnClickListener() {
                            @Override
                            public void onClick(DialogInterface dialog, int which) {
                                final Context context = getContext();
                                if (getSheetCount() > 1) {
                                    deleteSheet(currentSheetNum);
                                    Toast.makeText(context, context.getText(R.string.message_sheet_delete_success), Toast.LENGTH_LONG).show();  //**string.xml?
                                } else {
                                    Toast.makeText(context, context.getText(R.string.message_sheet_delete_at_least_one), Toast.LENGTH_LONG).show();  //**string.xml?
                                }
                                dialog.dismiss();     // 닫기
                            }
                        })
                        .setNegativeButton(getContext().getString(R.string.button_no), new DialogInterface.OnClickListener() {
                            @Override
                            public void onClick(DialogInterface dialog, int which) {
                                dialog.dismiss();     // 닫기
                            }
                        })
                        .create();
        }

        return null;
    }

    private void showSoftKeyboard() {
        requestFocus();
        inputMethodManager.showSoftInput(this, 0);
    }

    private void hideSoftKeyboard() {
        inputMethodManager.hideSoftInputFromWindow(getWindowToken(), 0);
    }

    @Override
    public boolean onTouchEvent(MotionEvent event) {
        super.onTouchEvent(event);

        scaleGestureDetector.onTouchEvent(event);
        gestureDetector.onTouchEvent(event);

        // scroll의 up 이벤트는 여기서 감지합니다.
        if (event.getAction() == MotionEvent.ACTION_UP && scrollFlag && selectionFlag) {
            if (DEBUG) {
                Log.d(TAG, "multiple cell selection : (" + selectedStartColumn + ", " + selectedStartRow + ")(" + selectedEndColumn + ", " + selectedEndRow + ")");
            }
            scrollFlag = false;
            selectionFlag = false;
        }

        return true;
    }

    private Cell getCell(int sheetNum, int column, int row) {
        if (workBookController != null) {
            return workBookController.getCell(sheetNum, column, row);
        } else {
            int[] allSheetId = workBook.getAllSheetId();
            Sheet sheet = workBook.getSheet(allSheetId[sheetNum]);

            return sheet.getCell(column, row);
        }
    }

    private Cell setCell(int sheetNum, int column, int row, String cellValue) {
        int cellType = validateValue(cellValue);

        if (workBookController != null) {
            return workBookController.setCell(sheetNum, column, row, cellValue, cellType);
        } else {
            int[] allSheetId = workBook.getAllSheetId();
            Sheet sheet = workBook.getSheet(allSheetId[sheetNum]);

            return sheet.createCell(column, row, cellValue, cellType);
        }
    }

    private int validateValue(String value) {

        if (value.startsWith(EXCEL_FUNCTION_START_CHAR)) {
            if (ExcelFunctionFactory.isExcelFunction(value)) {
                return TYPE_CELL_FUNCTION;
            } else {
                return TYPE_CELL_STRING;
            }
        } else if (Utils.isNumber(value)) {
            return TYPE_CELL_NUMBER;
        } else {
            return TYPE_CELL_STRING;
        }
    }

    private String getSheetName(int sheetNum) {
        if (workBookController != null) {
            return workBookController.getSheetName(sheetNum);
        } else {
            int[] allSheetId = workBook.getAllSheetId();
            Sheet sheet = workBook.getSheet(allSheetId[sheetNum]);

            return sheet.getName();
        }
    }

    private void setSheetName(int sheetNum, String newSheetName) {
        if (workBookController != null) {
            workBookController.setSheetName(sheetNum, newSheetName);
        } else {
            int[] allSheetId = workBook.getAllSheetId();

            workBook.setSheetName(allSheetId[sheetNum], newSheetName);
        }
    }

    private int getSheetCount() {
        if (workBookController != null) {
            return workBookController.getSheetCount();
        } else {
            return workBook.getAllSheetId().length;
        }
    }

    private Sheet addSheet() {
        if (workBookController != null) {
            return workBookController.addSheet();
        } else {
            return workBook.createSheet();
        }
    }

    private void deleteSheet(int sheetNum) {
        if (workBookController != null) {
            workBookController.deleteSheet(sheetNum);
        } else {
            int[] allSheetId = workBook.getAllSheetId();

            workBook.deleteSheet(allSheetId[sheetNum]);
        }
    }

    private void changeSheetOrder(int sheetNum, int newSheetOrder) {
        if (workBookController != null) {
            workBookController.changeSheetOrder(sheetNum, newSheetOrder);
        } else {
            int[] allSheetId = workBook.getAllSheetId();

            workBook.changeSheetOrder(allSheetId[sheetNum], newSheetOrder);
        }
    }
}