package com.study.hancom.myexcel.util;

import android.content.ContentResolver;
import android.content.Context;
import android.net.Uri;
import android.os.Environment;
import android.util.Log;

import com.study.hancom.myexcel.model.Cell;
import com.study.hancom.myexcel.model.Sheet;
import com.study.hancom.myexcel.model.WorkBook;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.study.hancom.myexcel.BuildConfig.DEBUG;
import static com.study.hancom.myexcel.model.WorkBook.DEFAULT_SHEET_MAX_COLUMN;
import static com.study.hancom.myexcel.model.WorkBook.DEFAULT_SHEET_MAX_ROW;

public final class Hana {
    private static final String TAG = "Hana";

    public static final String HANA_FILENAME_EXTENSION = ".hana";

    private static final String SHEET_BRACKET_OPENER = "{";
    private static final String SHEET_BRACKET_CLOSER = "}";

    private static final String SECTION_SEPARATOR = ":";
    private static final String HEADER_SECTION_NAME = "header";
    private static final String CONTENT_SECTION_NAME = "content";

    private static final String CELL_BRACKET_OPENER = "{";
    private static final String CELL_BRACKET_CLOSER = "}";

    private static final String ATTRIBUTE_SEPARATOR = "=";
    private static final String HEADER_ATTRIBUTE_CELLSEPARATOR_NAME = "cellSeparator";
    private static final String HEADER_ATTRIBUTE_FIRSTROW_NAME = "firstRow";
    private static final String HEADER_ATTRIBUTE_FIRSTCOLUMN_NAME = "firstColumn";

    private static final String CELL_VALUE_TYPE_SEPARATOR = ",";
    private static final String CELL_VALUE_SEPARATOR = "\"";

    private static final String DEFAULT_CELL_SEPARATOR = ",";

    private String cellSeparator = null;
    private int firstColumn = 1;
    private int firstRow = 1;

    //**에러는 일단 던집니다.
    public String encodeFile(String fileName, WorkBook workBook) throws Exception {

        final String fileNameWithExtension = fileName + HANA_FILENAME_EXTENSION;
        final File directory = new File(Environment.getExternalStorageDirectory().getAbsolutePath() + "/MyExcel");

        // 디렉토리를 생성합니다.
        directory.mkdirs();

        final File file = new File(directory, fileNameWithExtension);

        if (!file.createNewFile()) {    // 앱의 저장공간 접근 권한이 허용되어 있어야 합니다.
            if (DEBUG) {
                Log.d(TAG, "File already exists");
            }
        }

        if (file.exists()) {
            FileOutputStream outputStream = null;

            try {
                outputStream = new FileOutputStream(file);
                StringBuilder content = new StringBuilder();

                for (int eachSheetId : workBook.getAllSheetId()) {
                    Sheet sheet = workBook.getSheet(eachSheetId);

                    content.append(sheet.getName());
                    content.append(SHEET_BRACKET_OPENER);
                    linebreak(content);

                    /* 헤더 섹션 */
                    content.append(HEADER_SECTION_NAME);
                    content.append(SECTION_SEPARATOR);
                    linebreak(content);

                    content.append(HEADER_ATTRIBUTE_CELLSEPARATOR_NAME);
                    content.append(ATTRIBUTE_SEPARATOR);
                    content.append(DEFAULT_CELL_SEPARATOR);
                    linebreak(content);

                    int maxColumn = DEFAULT_SHEET_MAX_COLUMN;
                    int maxRow = DEFAULT_SHEET_MAX_ROW;
                    int firstColumn = maxColumn;
                    int firstRow = maxRow;
                    int lastColumn = 1;
                    int lastRow = 1;

                    for (int i = 1; i <= maxRow; i++) {
                        for (int j = 1; j <= maxColumn; j++) {
                            if (sheet.getCell(j, i) != null) {
                                if (firstColumn > j) {
                                    firstColumn = j;
                                }
                                if (firstRow > i) {
                                    firstRow = i;
                                }
                                if (lastColumn < j) {
                                    lastColumn = j;
                                }
                                if (lastRow < i) {
                                    lastRow = i;
                                }
                            }
                        }
                    }

                    content.append(HEADER_ATTRIBUTE_FIRSTROW_NAME);
                    content.append(ATTRIBUTE_SEPARATOR);
                    content.append(firstRow);
                    linebreak(content);

                    content.append(HEADER_ATTRIBUTE_FIRSTCOLUMN_NAME);
                    content.append(ATTRIBUTE_SEPARATOR);
                    content.append(firstColumn);
                    linebreak(content, 2);

                    /* 콘텐트 섹션 */
                    content.append(CONTENT_SECTION_NAME);
                    content.append(SECTION_SEPARATOR);
                    linebreak(content);

                    for (int row = firstRow; row <= lastRow; row++) {
                        for (int column = firstColumn; column <= lastColumn; column++) {
                            Cell cell = sheet.getCell(column, row);
                            if (cell != null) {
                                String cellValue = cell.getValue();
                                content.append(CELL_BRACKET_OPENER);
                                if (cellValue.contains(CELL_BRACKET_OPENER) || cellValue.contains(CELL_BRACKET_CLOSER)
                                        || cellValue.contains(CELL_VALUE_TYPE_SEPARATOR) || cellValue.contains(CELL_VALUE_SEPARATOR)) {
                                    content.append(CELL_VALUE_SEPARATOR);
                                    content.append(cellValue);
                                    content.append(CELL_VALUE_SEPARATOR);
                                } else {
                                    content.append(cellValue);
                                }
                                content.append(CELL_VALUE_TYPE_SEPARATOR);
                                content.append(cell.getType());
                                content.append(CELL_BRACKET_CLOSER);
                            }
                            content.append(DEFAULT_CELL_SEPARATOR);
                        }
                        linebreak(content);
                    }

                    content.append(SHEET_BRACKET_CLOSER);
                    linebreak(content, 2);
                }

                outputStream.write(content.toString().getBytes(Charset.forName("UTF-8")));
                outputStream.flush();
            } finally {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
        }

        return fileNameWithExtension;
    }

    //**유틸
    public static void linebreak(StringBuilder stringBuilder) {
        linebreak(stringBuilder, 1);
    }

    public static void linebreak(StringBuilder stringBuilder, int count) {
        for (int i = 0; i < count; i++) {
            stringBuilder.append(System.getProperty("line.separator"));
        }
    }

    public WorkBook decodeFile(Context context, Uri uri) throws Exception {

        WorkBook workBook = null;
        InputStream inputStream = null;

        try {
            // 파일을 읽기 위한 InputStream을 생성합니다.
            // ContentResolver를 통해 데이터에 접근합니다.
            ContentResolver contentResolver = context.getContentResolver();
            inputStream = contentResolver.openInputStream(uri);
            BufferedReader bufferedReader = null;

            try {
                bufferedReader = new BufferedReader(new InputStreamReader(inputStream, "utf-8"));

                // 데이터 모델을 생성합니다.
                workBook = new WorkBook();
                Sheet sheet = null;

                boolean insideSheet = false;

                int readColumn;
                int readRow = 0;

                String eachLine;

                // header 영역과 content 영역을 구분하여 처리합니다.
                while ((eachLine = bufferedReader.readLine()) != null) {
                    String sheetName;

                    if (eachLine.equals(SHEET_BRACKET_CLOSER)) {
                        readRow = 0;
                        insideSheet = false;
                    } else if (!insideSheet) {
                        // sheet 이름을 추출합니다.
                        int sheetNameSeparatorIndex = eachLine.indexOf(SHEET_BRACKET_OPENER);

                        if (sheetNameSeparatorIndex > SHEET_BRACKET_OPENER.length()) {
                            insideSheet = true;
                            sheetName = eachLine.substring(0, sheetNameSeparatorIndex);
                            // sheet를 생성합니다.
                            sheet = workBook.createSheet(sheetName);

                            if (DEBUG) {
                                Log.d(TAG, "sheet name : " + sheetName);
                            }
                        }
                    } else if (insideSheet) {
                        // header 영역을 처리합니다.
                        if (eachLine.startsWith(HEADER_SECTION_NAME)) {
                            String attributeName, attributeValue;
                            int attributeSeparatorPosition;

                            // content 영역을 만나기 전까지 header 영역으로 간주합니다.
                            while (!eachLine.startsWith(CONTENT_SECTION_NAME)) {
                                // 속성 이름과 값을 구분자로 분리합니다.
                                // 속성 이름의 대소문자는 구분하지 않습니다.
                                attributeSeparatorPosition = eachLine.indexOf(ATTRIBUTE_SEPARATOR);

                                // 구분자를 찾지 못한 줄은 무시합니다.
                                if (attributeSeparatorPosition > 0) {
                                    attributeName = eachLine.substring(0, attributeSeparatorPosition);
                                    attributeValue = eachLine.substring(attributeSeparatorPosition + 1, eachLine.length());

                                    // 속성 이름에 따라 값을 저장합니다.
                                    if (attributeName.equals(HEADER_ATTRIBUTE_CELLSEPARATOR_NAME)) {
                                        cellSeparator = attributeValue;
                                    } else if (attributeName.equals(HEADER_ATTRIBUTE_FIRSTCOLUMN_NAME)) {
                                        firstColumn = Integer.parseInt(attributeValue);
                                    } else if (attributeName.equals(HEADER_ATTRIBUTE_FIRSTROW_NAME)) {
                                        firstRow = Integer.parseInt(attributeValue);
                                    }
                                }
                                eachLine = bufferedReader.readLine();
                            }

                            if (DEBUG) {
                                Log.d(TAG, "cellSeparator : " + cellSeparator);
                                Log.d(TAG, "firstColumn : " + firstColumn);
                                Log.d(TAG, "firstRow : " + firstRow);
                            }
                        } else {
                            // content 영역을 처리합니다.

                            // 정규식을 이용해 cell 단위로 데이터를 추출합니다.
                            // 아래는 "{value, type}" 형태의 문자열, 즉 cell 단위로 문자열을 검출하기 위한 정규식입니다.
                            final Pattern regex = Pattern.compile("(\\{((.*?)(,[0-9]+))\\})?(,)?");
                            final Matcher cells = regex.matcher(eachLine);

                            readColumn = 0;

                            while (cells.find()) {
                                String eachCell = cells.group();
                                int cellLength = eachCell.length();

                                if (cellLength > cellSeparator.length()) {
                                    int endIndex = cellLength - 1;

                                    // 대괄호 {}와 cell seperator를 걸러냅니다.
                                    if (eachCell.endsWith(",")) {
                                        --endIndex;
                                    }

                                    eachCell = eachCell.substring(1, endIndex);

                                    // value와 type을 분리합니다.
                                    final int cellValueTypeSeparatorIndex = eachCell.lastIndexOf(CELL_VALUE_TYPE_SEPARATOR);
                                    final int type = Integer.parseInt(eachCell.substring(cellValueTypeSeparatorIndex + 1, eachCell.length()));
                                    String value = eachCell.substring(0, cellValueTypeSeparatorIndex);

                                    if (value.endsWith(String.valueOf(CELL_VALUE_SEPARATOR))) {
                                        value = value.substring(1, value.length() - 1);
                                    }
                                    if (DEBUG) {
                                        Log.d(TAG, "(" + (firstColumn + readColumn) + ", " + (firstRow + readRow) + ")" + "[" + value + ", " + type + "]");
                                    }

                                    // sheet에 cell을 생성합니다.
                                    sheet.createCell((firstColumn + readColumn), (firstRow + readRow), value, type);
                                }
                                readColumn++;
                            }
                            readRow++;
                        }
                    }
                }
            } finally {
                if (bufferedReader != null) {
                    bufferedReader.close();
                }
            }
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
        }

        return workBook;
    }
}