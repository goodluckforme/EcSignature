package com.example.readpoi;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.content.Intent;
import android.view.Menu;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.Toast;

import com.example.signature.RxFileTool;
import com.example.signature.RxZipTool;
import com.example.wifitransfer.Constant;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class MainActivity extends Activity {
    public String filePath = Environment.getExternalStorageDirectory()
            + "/aa.xlsx";

    private Button btn_load;
    private Button btn_load_ppt;
    private Button btn_load_list;
    private Button btn_wifi;
    private Button btn_clear;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        btn_clear = (Button) findViewById(R.id.btn_clear);
        btn_load = (Button) findViewById(R.id.btn_load);
        btn_load_ppt = (Button) findViewById(R.id.btn_load_ppt);
        btn_load_list = (Button) findViewById(R.id.btn_load_list);
        btn_wifi = (Button) findViewById(R.id.btn_wifi);
        btn_load.setOnClickListener(new OnClickListener() {

            @Override
            public void onClick(View v) {
                Intent i = new Intent();
                i.setClass(MainActivity.this, FileReadActivity.class);
                Bundle bundle = new Bundle();
                filePath = Environment.getExternalStorageDirectory()
                        + "/bb.docx";
                bundle.putString("filePath", filePath);
                i.putExtras(bundle);
                startActivity(i);
            }
        });
        btn_load_ppt.setOnClickListener(new OnClickListener() {

            @Override
            public void onClick(View v) {
                Intent i = new Intent();
                i.setClass(MainActivity.this, ReadPPTX.class);
                startActivity(i);
            }
        });
        btn_load_list.setOnClickListener(new OnClickListener() {

            @Override
            public void onClick(View v) {
                Intent i = new Intent(MainActivity.this, ReadFActivity.class);
                i.setFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP | Intent.FLAG_ACTIVITY_SINGLE_TOP);
                startActivity(i);
            }
        });
        btn_wifi.setOnClickListener(new OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent i = new Intent();
                i.setClass(MainActivity.this, WifiBookActivity.class);
                startActivity(i);
            }
        });
        btn_clear.setOnClickListener(new OnClickListener() {
            @Override
            public void onClick(View v) {
                boolean dir = RxFileTool.deleteDir(Constant.PATH_DATA);
                boolean cache = RxFileTool.cleanExternalCache(MainActivity.this);
                if (dir && cache) {
                    Toast.makeText(MainActivity.this, "清空成功", Toast.LENGTH_SHORT).show();
                    initDatas();
                } else {
                    Toast.makeText(MainActivity.this, "清空失败", Toast.LENGTH_SHORT).show();
                }
            }
        });
        initDatas();
    }

    private void initDatas() {
        try {
            File file = new File("docx.zip");
            File destinationFile = new File(Environment.getExternalStorageDirectory() + File.separator + "piopic");
            File result = new File(destinationFile, file.getName());
            if (!result.exists()) {
                InputStream source = getAssets().open(file.getPath());
                if (RxFileTool.writeFileFromIS(result, source, false)) {
                    //解压出文件夹
                    RxZipTool.unzipFileByKeyword(result, destinationFile, "123456");
                } else {
                    Toast.makeText(MainActivity.this, "初始化失败", Toast.LENGTH_SHORT).show();
                }
            } else {
                //解压出文件夹
                RxZipTool.unzipFileByKeyword(result, destinationFile, "123456");
            }
        } catch (IOException e) {
            Toast.makeText(MainActivity.this, "初始化失败--" + e.getLocalizedMessage(), Toast.LENGTH_SHORT).show();
            e.printStackTrace();
        }
    }
}
