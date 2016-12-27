package com.study.hancom.myexcel.activity;

import android.content.Context;
import android.content.DialogInterface;
import android.content.Intent;
import android.net.Uri;
import android.support.v7.app.AlertDialog;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import com.study.hancom.myexcel.R;
import com.study.hancom.myexcel.controller.WorkBookController;
import com.study.hancom.myexcel.view.WorkBookView;

import static com.study.hancom.myexcel.BuildConfig.DEBUG;
import static com.study.hancom.myexcel.util.Hana.HANA_FILENAME_EXTENSION;
import static com.study.hancom.myexcel.util.listener.DataChangeListenerService.setDataChangeListener;

public class MainActivity extends AppCompatActivity {

    private static final int REQUEST_CODE_LOAD_FILE = 1;

    private WorkBookController workBookController = new WorkBookController();
    private WorkBookView workBookView = null;

    private Button newFileButton = null;
    private Button loadFileButton = null;
    private Button saveFileButton = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        setContentView(R.layout.activity_main);

        workBookView = (WorkBookView) findViewById(R.id.workBookView);
        workBookView.setController(workBookController);
        setDataChangeListener(workBookView);

        newFileButton = (Button) findViewById(R.id.button_new);
        newFileButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (workBookView != null) {
                    workBookController.init();
                    Toast.makeText(MainActivity.this, getString(R.string.message_new_document), Toast.LENGTH_LONG).show();
                }
            }
        });

        loadFileButton = (Button) findViewById(R.id.button_load);
        // 파일 불러오기 버튼의 이벤트를 정의합니다.
        // 인텐트로 탐색기를 열어 파일을 엽니다.
        loadFileButton.setOnClickListener(new Button.OnClickListener() {
            public void onClick(View v) {
                if (workBookView != null) {
                    //**intent-filter?
                    Intent intent = new Intent(Intent.ACTION_GET_CONTENT);

                    intent.setType("*/*");
                    startActivityForResult(intent, REQUEST_CODE_LOAD_FILE);
                }
            }
        });

        saveFileButton = (Button) findViewById(R.id.button_save);
        // 파일 저장하기 버튼의 이벤트를 정의합니다.
        // 인텐트로 탐색기를 열어 파일을 생성할 경로를 얻어옵니다.
        saveFileButton.setOnClickListener(new Button.OnClickListener() {
            public void onClick(View v) {
                if (workBookView != null) {
                    final Context context = MainActivity.this;
                    final EditText editFileName = new EditText(context);
                    String fileName = workBookController.getFileName();

                    if (fileName != null) {
                        String fileNameWithoutExtension = fileName.substring(0, fileName.length() - HANA_FILENAME_EXTENSION.length());

                        editFileName.setText(fileNameWithoutExtension);
                        editFileName.setSelection(fileNameWithoutExtension.length());
                    }

                    new AlertDialog.Builder(context)
                            .setTitle(getString(R.string.button_save_file))
                            .setMessage(getString(R.string.message_file_name))
                            .setView(editFileName)
                            .setPositiveButton(getString(R.string.button_ok), new DialogInterface.OnClickListener() {
                                @Override
                                public void onClick(DialogInterface dialog, int which) {
                                    String inputSheetName = editFileName.getText().toString();
                                    if (inputSheetName.length() > 0) {
                                        try {
                                            workBookController.saveHana(context, inputSheetName);
                                            Toast.makeText(context, getString(R.string.message_save_file_success), Toast.LENGTH_LONG).show();  //**temp
                                        } catch (Exception e) {
                                            if (DEBUG) {
                                                e.printStackTrace();
                                            }
                                            Toast.makeText(context, getString(R.string.message_save_file_fail), Toast.LENGTH_LONG).show();  //**temp
                                        } finally {
                                            dialog.dismiss();     // 닫기
                                        }
                                    } else {
                                        Toast.makeText(context, getString(R.string.message_save_file_name_too_short), Toast.LENGTH_LONG).show();  //**string.xml?
                                    }
                                }
                            })
                            .setNegativeButton(context.getString(R.string.button_cancel), new DialogInterface.OnClickListener() {
                                @Override
                                public void onClick(DialogInterface dialog, int which) {
                                    dialog.dismiss();     // 닫기
                                }
                            })
                            .create().show();
                }
            }
        });
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        switch (requestCode) {
            case REQUEST_CODE_LOAD_FILE:
                if (resultCode == RESULT_OK) {
                    Uri uri = data.getData();
                    try {
                        workBookController.readHana(this, uri);
                        Toast.makeText(this, getString(R.string.message_load_file_success), Toast.LENGTH_LONG).show();
                    } catch (Exception e) {
                        e.printStackTrace();
                        Toast.makeText(this, getString(R.string.message_load_file_fail), Toast.LENGTH_LONG).show();
                    }
                }

                break;
        }

    }
}