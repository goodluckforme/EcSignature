package com.example.readpoi;

import android.Manifest;
import android.annotation.TargetApi;
import android.app.Activity;
import android.content.DialogInterface;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.os.StrictMode;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.util.Xml;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.BaseAdapter;
import android.widget.ListView;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

//import org.textmining.text.extraction.WordExtractor;

//import jxl.Cell;
//import jxl.CellType;
//import jxl.DateCell;
//import jxl.NumberCell;
//import jxl.Sheet;
//import jxl.Workbook;


public class ReadFActivity extends Activity {
    private TextView tv = null;
    private ListView list = null;
    private List<String> paths = new ArrayList<>();//存放路径
    private List<String> items = new ArrayList<>();//存放名称
    private String rootPath = Environment.getExternalStorageDirectory() + "";

    private String resultDate;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_read_f);
        tv = findViewById(R.id.R_tv);
        list = findViewById(R.id.listq);
        getPermission();
    }

    void getPermission() {
        int permissionCheck1 = ContextCompat.checkSelfPermission(getApplicationContext(), Manifest.permission.READ_EXTERNAL_STORAGE);
        int permissionCheck2 = ContextCompat.checkSelfPermission(getApplicationContext(), Manifest.permission.WRITE_EXTERNAL_STORAGE);
        if (permissionCheck1 != PackageManager.PERMISSION_GRANTED || permissionCheck2 != PackageManager.PERMISSION_GRANTED) {
            ActivityCompat.requestPermissions(this,
                    new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE, Manifest.permission.READ_EXTERNAL_STORAGE},
                    124);
        } else
            getFileDir(rootPath);
    }

    @Override
    public void onRequestPermissionsResult(int requestCode,
                                           String[] permissions,
                                           int[] grantResults) {
        if (requestCode == 124) {
            if ((grantResults.length > 0) && (grantResults[0] == PackageManager.PERMISSION_GRANTED)) {
                //获取rootPath目录下的文件.
                getFileDir(rootPath);
            } else {
                Toast.makeText(this, "授权失败", Toast.LENGTH_LONG).show();
            }
        }
    }

    public void getFileDir(String filePath) {
        try {
            this.tv.setText("当前路径:" + filePath);// 设置当前所在路径
            items = new ArrayList<>();
            paths = new ArrayList<>();
            File f = new File(filePath);
            if (!f.isDirectory()) return;
            File[] files = f.listFiles();// 列出所有文件
//            // 如果不是根目录,则列出返回根目录和上一目录选项
            if (!filePath.equals(rootPath)) {
                items.add("返回根目录");
                paths.add(rootPath);
                items.add("返回上一层目录");
                paths.add(f.getParent());
            }
            // 将所有文件存入list中
            if (files != null) {
                int count = files.length;// 文件个数
                for (int i = 0; i < count; i++) {
                    File file = files[i];
                    String name = file.getName();
                    if (name.endsWith(".doc") ||
                            name.endsWith(".docx") ||
                            name.endsWith(".xsl") ||
                            name.endsWith(".xlsx") ||
                            name.endsWith(".pptx") || file.isDirectory()) {
                        items.add(name);
                        paths.add(file.getAbsolutePath());
                    }
                }
            }
            ArrayAdapter<String> adapter = new ArrayAdapter<>(this, android.R.layout.simple_list_item_1, items);
            list.setAdapter(adapter);
            list.setOnItemClickListener(new AdapterView.OnItemClickListener() {
                @Override
                public void onItemClick(AdapterView<?> parent, View view, int position, long id) {
                    String path = paths.get(position);
                    String name = items.get(position);
                    if ("返回根目录".equals(name)) {
                        getFileDir(rootPath);
                    } else if ("返回上一层目录".equals(name)) {
                        getFileDir(path);
                    } else {
                        File file = new File(path);
                        if (file.isDirectory()) {
                            getFileDir(path);
                            tv.setText(path);
                        } else {
                            Intent i = new Intent();
                            i.setClass(ReadFActivity.this, FileReadActivity.class);
                            Bundle bundle = new Bundle();
                            bundle.putString("filePath", path);
                            i.putExtras(bundle);
                            startActivity(i);
                        }
                    }
                }
            });
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    @Override
    protected void onResume() {
        super.onResume();
        if (null != list&&null != list.getAdapter()) {
            getWord(null);
        }
    }

    public void getWord(View view) {
        items.clear();
        paths.clear();
        items.add("返回根目录");
        paths.add(rootPath);
        getAllFiles(rootPath + "/piopic/docx");
        ((BaseAdapter) list.getAdapter()).notifyDataSetChanged();
        this.tv.setText("当前路径:" + rootPath + "/piopic/docx");// 设置当前所在路径
    }

    private void getAllFiles(String newPath) {
        File file = new File(newPath);
        if (file.isDirectory()) {
            for (File item : file.listFiles()) {
                if (item.isDirectory()) {
                    getAllFiles(newPath);
                } else {
                    String name = item.getName();
                    if (name.endsWith(".doc") ||
                            name.endsWith(".docx")) {
                        items.add(name);
                        paths.add(item.getAbsolutePath());
                    }
                }
            }
        } else {
            if (newPath.endsWith(".doc") ||
                    newPath.endsWith(".docx")) {
                items.add(file.getName());
                paths.add(newPath);
            }
        }
    }

    public String readWord(String file) {
        // 创建输入流读取doc文件
        FileInputStream in = null;
        String text = null;
        try {
            in = new FileInputStream(new File(file));
            // in=this.openFileInput("aaa.doc");
            WordExtractor extractor = null;
            // 创建WordExtractor
            extractor = new WordExtractor(in);
            // 对doc文件进行提取
            text = extractor.getText();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return text;
    }

//    public String readXLS(String path) {
//        String str = "";
//        try {
//            Workbook workbook = null;
//            workbook = Workbook.getWorkbook(new File(path));
//            Sheet sheet = workbook.getSheet(0);
//            Cell cell = null;
//            int columnCount = sheet.getColumns();
//            int rowCount = sheet.getRows();
//            for (int i = 0; i < rowCount; i++) {
//                for (int j = 0; j < columnCount; j++) {
//                    cell = sheet.getCell(j, i);
//                    String temp2 = "";
//                    if (cell.getType() == CellType.NUMBER) {
//                        temp2 = ((NumberCell) cell).getValue() + "";
//                    } else if (cell.getType() == CellType.DATE) {
//                        temp2 = "" + ((DateCell) cell).getDate();
//                    } else {
//                        temp2 = "" + cell.getContents();
//                    }
//                    str = str + "  " + temp2;
//                }
//                str += " ";
//            }
//            workbook.close();
//        } catch (Exception e) {
//        }
//        if (str == null) {
//            str = "解析文件出现问题";
//        }
//        return str;
//
//    }

    public static String readDOCX(String path) {
        String river = "";
        try {
            ZipFile xlsxFile = new ZipFile(new File(path));
            ZipEntry sharedStringXML = xlsxFile.getEntry("word/document.xml");
            InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
            XmlPullParser xmlParser = Xml.newPullParser();
            xmlParser.setInput(inputStream, "utf-8");
            int evtType = xmlParser.getEventType();
            while (evtType != XmlPullParser.END_DOCUMENT) {
                switch (evtType) {
                    case XmlPullParser.START_TAG:
                        String tag = xmlParser.getName();
                        System.out.println(tag);
                        if (tag.equalsIgnoreCase("t")) {
                            river += xmlParser.nextText() + " ";
                        }
                        break;
                    case XmlPullParser.END_TAG:
                        break;
                    default:
                        break;
                }
                evtType = xmlParser.next();
            }
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlPullParserException e) {
            e.printStackTrace();
        }
        if (river == null) {
            river = "解析文件出现问题";
        }
        return river;

    }


    class onclick1 implements DialogInterface.OnClickListener {


        @TargetApi(Build.VERSION_CODES.HONEYCOMB)
        @Override
        public void onClick(DialogInterface dialog, int option) {
            switch (option) {
                case 0:
                    StrictMode.setThreadPolicy(new
                            StrictMode.ThreadPolicy
                                    .Builder().detectDiskReads().detectDiskWrites().detectNetwork
                            ().penaltyLog().build());
                    StrictMode.setVmPolicy(new
                            StrictMode.VmPolicy.Builder
                            ().detectLeakedSqlLiteObjects().detectLeakedClosableObjects
                            ().penaltyLog().penaltyDeath().build());
                    String httpUrl
                            = "http://192.168.0.145/aa.txt";

                    HttpURLConnection httpConnection;
                    URL url;
                    int code;
                    try {
                        url = new URL(httpUrl);
                        URLConnection connection = url.openConnection();
                        httpConnection = (HttpURLConnection) connection;
                        httpConnection = (HttpURLConnection) url.openConnection();
                        httpConnection.setRequestMethod("GET");
                        InputStreamReader in = new
                                InputStreamReader(httpConnection.getInputStream());
                        BufferedReader
                                buffer = new BufferedReader(in);
                        String inputLine = null;
                        while ((inputLine = buffer.readLine()) != null) {
                            resultDate += inputLine + "\n";
                            tv.setText("aaaa");


                        }
                        in.close();
                        httpConnection.disconnect();
                        if (resultDate != null) {

                            tv.setText(resultDate);
                        }
                    } catch (Exception e) {
                    }

            }

        }
    }
}
