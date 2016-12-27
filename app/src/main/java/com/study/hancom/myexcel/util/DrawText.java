package com.study.hancom.myexcel.util;

import android.graphics.Canvas;
import android.graphics.Paint;

public class DrawText {

    private static final String ELLIPSIS_STRING = "";

    public static final int ALIGNMENT_LEFT = 0;
    public static final int ALIGNMENT_RIGHT = 1;
    public static final int ALIGNMENT_CENTER = 2;

    public static void drawText(String text, int alignment, float x, float y, float width, float height, Paint paint, Canvas canvas, boolean ellipsis) {
        // 텍스트를 그립니다.

        if (text.length() > 0 && width > 0) {
            final Paint.FontMetrics fontMetrics = paint.getFontMetrics();
            float textHeight = (fontMetrics.top * -1) - fontMetrics.bottom;
            float textWidth = paint.measureText(text);
            String resultText = text;

            // 말줄임표(ellipsis) 옵션에 대해 true면
            // 출력할 공간의 너비를 넘지 않도록 텍스트를 잘라내고 말줄임표를 붙입니다.
            if (textWidth > width && ellipsis) {
                final float ellipsisWidth = paint.measureText(ELLIPSIS_STRING);
                float charsWidth = 0;

                for (int i = 0; i < text.length(); i++) {
                    charsWidth += paint.measureText(text, i, i + 1);
                    if (charsWidth > width - ellipsisWidth) {
                        resultText = text.substring(0, i);
                        resultText += ELLIPSIS_STRING;
                        break;
                    }
                }

                textWidth = paint.measureText(resultText);
            }

            final float fontX = (width - textWidth) / 2;
            final float fontY = (height + textHeight) / 2;

            switch (alignment) {
                // 왼쪽 정렬
                case ALIGNMENT_LEFT:
                    canvas.drawText(resultText, x, y + fontY, paint);

                    break;
                // 오른쪽 정렬
                case ALIGNMENT_RIGHT:
                    canvas.drawText(resultText, x + (fontX * 2), y + fontY, paint);

                    break;
                // 중앙 정렬 (기본)
                case ALIGNMENT_CENTER:
                    // no break ; default
                default:
                    canvas.drawText(resultText, x + fontX, y + fontY, paint);
                    break;
            }
        }
    }
}