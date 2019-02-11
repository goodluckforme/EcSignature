package com.example.readpoi;

import android.annotation.SuppressLint;
import android.app.Activity;
import android.app.AlertDialog;
import android.app.ProgressDialog;
import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.drawable.ColorDrawable;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.os.Handler;
import android.os.Looper;
import android.os.StrictMode;
import android.provider.MediaStore;
import android.util.Log;
import android.view.Gravity;
import android.view.View;
import android.webkit.JavascriptInterface;
import android.webkit.ValueCallback;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.ImageView;
import android.widget.LinearLayout;
import android.widget.PopupWindow;
import android.widget.TextView;
import android.widget.Toast;

import com.example.signature.RxFileTool;
import com.example.signature.RxZipTool;
import com.example.signature.SignatureView;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.TableIterator;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;


public class FileReadActivity extends Activity {
    public Range range = null;
    public HWPFDocument hwpf = null;
    public String htmlPath;
    public String picturePath;
    public WebView webView;
    public ImageView sign;
    public List pictures;
    public TableIterator tableIterator;
    public int presentPicture = 0;
    public int screenWidth;
    public FileOutputStream output;
    public File myFile;
    StringBuffer lsb = new StringBuffer();
    FR fr = null;
    TextView tv;
    private View document;
    private Handler handler;
    private String lasSignFilePath = "/storage/emulated/0/piopic/Signature_1547006052936.jpg";
    private boolean root;
    private String returnPath;
    private boolean onPageFinished;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.load_view_activity);
        webView = findViewById(R.id.wv_view);
        document = findViewById(R.id.document);
        sign = findViewById(R.id.sign);
        handler = new Handler(Looper.getMainLooper());

        WebSettings webSettings = webView.getSettings();
        webSettings.setBuiltInZoomControls(false);
        webSettings.setUseWideViewPort(false);
        webSettings.setSupportZoom(false);
        // 设置与Js交互的权限
        webSettings.setJavaScriptEnabled(true);
        // 设置允许JS弹窗
        webSettings.setJavaScriptCanOpenWindowsAutomatically(true);

        // 通过addJavascriptInterface()将Java对象映射到JS对象
        //参数1：Javascript对象名
        //参数2：Java对象名
        webView.addJavascriptInterface(new AndroidtoJs(), "callAndroid");//AndroidtoJS类对象映射到js的test对象

        webView.setWebViewClient(new WebViewClient() {
            @Override
            public void onPageFinished(WebView view, String url) {
                Toast.makeText(FileReadActivity.this, "onPageFinished", Toast.LENGTH_SHORT).show();
                onPageFinished = true;
                super.onPageFinished(view, url);
            }
        });

        // 创建意图 获得要显示的文件
        Intent intent = this.getIntent();
        Bundle bundle = intent.getExtras();
        String nameStr = bundle.getString("filePath");
        String path = new FR(nameStr).returnPath;
        webView.loadUrl(path);
//        root = RxShellTool.isRoot();
//        if (root) {
//            Toast.makeText(FileReadActivity.this, "已经Root", Toast.LENGTH_SHORT).show();
//        } else {
//            Toast.makeText(FileReadActivity.this, "没有Root", Toast.LENGTH_SHORT).show();
//        }
        returnPath = RxFileTool.getPathFromUri(FileReadActivity.this, Uri.parse(path));
    }

    //清空文件夹
    public void clear(View view) {
        final ProgressDialog progressDialog = new ProgressDialog(FileReadActivity.this);
        progressDialog.setTitle("正在处理中，请稍后");
        progressDialog.setCancelable(false);
        if (RxFileTool.delAllFile(new File(returnPath).getParent()))
            progressDialog.setTitle("清除成功");
        else
            progressDialog.setTitle("清除失败");
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                progressDialog.dismiss();
            }
        }, 500);
    }

    String zipEncrypt;

    public class AndroidtoJs extends Object {
        // 定义JS需要调用的方法
        // 被JS调用的方法必须加入@JavascriptInterface注解
        @JavascriptInterface
        @SuppressLint("JavascriptInterface")
        public void showPad() {
            handler.post(new Runnable() {
                @Override
                public void run() {
                    sign(null);
                }
            });
        }

        @JavascriptInterface
        public void getSource(String html) {
            final ProgressDialog progressDialog = new ProgressDialog(FileReadActivity.this);
            progressDialog.setCancelable(false);
            progressDialog.setTitle("正在处理当中，请稍后..");
            progressDialog.show();
            Log.e("getSource", html);
            final String name = RxFileTool.getFileNameNoExtension(returnPath);
            html = html.replaceAll("/storage/emulated/0/piopic/" + name, "./");
            RxFileTool.write(returnPath, html);
            File parentFile = new File(returnPath).getParentFile();
            final String absolutePath = parentFile.getAbsolutePath();
            List<File> files = RxFileTool.listFilesInDir(absolutePath);
            boolean toZip = false;
            try {
                zipEncrypt = RxZipTool.zipEncrypt(absolutePath, absolutePath + ".zip", true, "123456");
                toZip = zipEncrypt != null;
//                toZip = RxZipTool.zipFiles(files, absolutePath + name + ".zip", "");
            } catch (Exception e) {
                progressDialog.dismiss();
                Toast.makeText(FileReadActivity.this, "文档处理异常，请重试", Toast.LENGTH_SHORT).show();
                e.printStackTrace();
            }
            if (toZip) {
                handler.postDelayed(new Runnable() {
                    @Override
                    public void run() {
                        progressDialog.dismiss();
                        Toast.makeText(FileReadActivity.this, "文档处理完成", Toast.LENGTH_SHORT).show();
                        shareFile(zipEncrypt);
                    }
                }, 500);
            }
        }
    }

    public void sign(View view) {
        final PopupWindow popupWindow = new PopupWindow(LinearLayout.LayoutParams.MATCH_PARENT, LinearLayout.LayoutParams.MATCH_PARENT);
        popupWindow.setFocusable(true);
        popupWindow.setOutsideTouchable(false);
        popupWindow.setBackgroundDrawable(new ColorDrawable());
        View inflate = View.inflate(this, R.layout.layou_sign, null);
        popupWindow.setContentView(inflate);
        popupWindow.showAtLocation(findViewById(android.R.id.content), Gravity.NO_GRAVITY, 0, 0);
        final SignatureView signatureView = inflate.findViewById(R.id.signatureView);
        inflate.findViewById(R.id.confirm).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                File lastFile = new File(lasSignFilePath);
                if (lastFile.exists()) {
                    boolean delete = lastFile.delete();
                    if (delete) {
//删除图片
                        Uri uri = MediaStore.Images.Media.EXTERNAL_CONTENT_URI;
                        String where = MediaStore.Images.Media.DATA + "='" + lasSignFilePath + "'";
                        getContentResolver().delete(uri, where, null);
                        Toast.makeText(FileReadActivity.this, "删除成功", Toast.LENGTH_SHORT).show();
                    } else {
                        Toast.makeText(FileReadActivity.this, "删除失败", Toast.LENGTH_SHORT).show();
                    }
                }

                Bitmap signatureBitmap = signatureView.getSignatureBitmap();
//sign.setImageBitmap(signatureBitmap);
                String fileName = String.format("Signature_%d.jpg", System.currentTimeMillis());
                File lastSign = new File(new File(returnPath).getParent(), fileName);
                if (addSignatureToGallery(signatureBitmap, lastSign)) {
                    setImgToWb(lastSign);
                    lasSignFilePath = lastSign.getAbsolutePath();
                } else {
                    Toast.makeText(FileReadActivity.this, "保存失败", Toast.LENGTH_SHORT).show();
                }
                popupWindow.dismiss();
            }
        });
        inflate.findViewById(R.id.clear).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                signatureView.clear();
            }
        });
    }


    private void setImgToWb(File photo) {
        // Android版本变量
        final int version = Build.VERSION.SDK_INT;
// 因为该方法在 Android 4.4 版本才可使用，所以使用时需进行版本判断
        if (version < 18) {
            webView.loadUrl("javascript:setImgToWb(" + "\"" + photo.getAbsolutePath() + "\"" + ")");
        } else {
            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT) {
                webView.evaluateJavascript("javascript:setImgToWb(" + "\"" + photo.getAbsolutePath() + "\"" + ")", new ValueCallback<String>() {
                    @Override
                    public void onReceiveValue(String value) {
                        //此处为 js 返回的结果
                    }
                });
            }
        }
    }


    public File getAlbumStorageDir(String albumName) {
        // Get the directory for the user's public pictures directory.
        File file = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_PICTURES), albumName);
        if (!file.mkdirs()) {
            Log.e("SignaturePad", "Directory not created");
        }
        return file;
    }

    public void saveBitmapToJPG(Bitmap bitmap, File photo) throws IOException {
        Bitmap newBitmap = Bitmap.createBitmap(bitmap.getWidth(), bitmap.getHeight(), Bitmap.Config.ARGB_8888);
        Canvas canvas = new Canvas(newBitmap);
        canvas.drawColor(Color.WHITE);
        canvas.drawBitmap(bitmap, 0, 0, null);
        OutputStream stream = new FileOutputStream(photo);
        newBitmap.compress(Bitmap.CompressFormat.JPEG, 80, stream);
        stream.close();
    }

    public boolean addSignatureToGallery(Bitmap signature, File photo) {
        boolean result = false;
        try {
            //File photo = new File(getAlbumStorageDir("SignaturePad"), String.format("Signature_%d.jpg", System.currentTimeMillis()));

            saveBitmapToJPG(signature, photo);
            Intent mediaScanIntent = new Intent(Intent.ACTION_MEDIA_SCANNER_SCAN_FILE);
            Uri contentUri = Uri.fromFile(photo);
            mediaScanIntent.setData(contentUri);
            sendBroadcast(mediaScanIntent);
            result = true;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }


    //============================================保存签名文档======================================
    //RxShellTool.CommandResult commandResult1 = RxShellTool.execCmd("adb pull /storage/emulated/0/piopic/file C:\\Users\\Administrator\\Desktop", root);
    public void save(View view) {
        webView.loadUrl("javascript:window.callAndroid.getSource(document.documentElement.outerHTML);void(0)");
    }

    /**
     * 分享文件
     *
     * @param
     * @param path
     */
    public void shareFile(String path) {
        checkFileUriExposure();
        Intent intent = new Intent(Intent.ACTION_SEND);
        intent.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
        intent.putExtra(Intent.EXTRA_STREAM, Uri.fromFile(new File(path)));  //传输图片或者文件 采用流的方式
        intent.setType("*/*");   //分享文件
        startActivity(Intent.createChooser(intent, "分享"));
    }

    /**
     * 分享前必须执行本代码，主要用于兼容SDK18以上的系统
     */
    private void checkFileUriExposure() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.JELLY_BEAN_MR2) {
            StrictMode.VmPolicy.Builder builder = new StrictMode.VmPolicy.Builder();
            StrictMode.setVmPolicy(builder.build());
            builder.detectFileUriExposure();
        }
    }

}
