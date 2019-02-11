package com.example.readpoi;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.DialogInterface;
import android.content.Intent;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.support.annotation.Nullable;
import android.text.TextUtils;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.TextView;

import com.example.wifitransfer.Defaults;
import com.example.wifitransfer.FileSendListener;
import com.example.wifitransfer.NetworkUtils;
import com.example.wifitransfer.ServerRunner;

/**
 * Created by Administrator on 2019/1/11.
 */

public class WifiBookActivity extends Activity implements FileSendListener {
    TextView mTvWifiName;
    TextView mTvWifiIp;
    TextView tvRetry;

    @Override
    protected void onCreate(@Nullable Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_wifi_book);
        mTvWifiName = findViewById(R.id.mTvWifiName);
        mTvWifiIp = findViewById(R.id.mTvWifiIp);
        tvRetry = findViewById(R.id.tvRetry);

        initDatas();
    }

    private void initDatas() {
        String wifiname = NetworkUtils.getConnectWifiSsid(this);
        if (!TextUtils.isEmpty(wifiname)) {
            mTvWifiName.setText(wifiname.replace("\"", ""));
        } else {
            mTvWifiName.setText("Unknow");
        }

        String wifiIp = NetworkUtils.getConnectWifiIp(this);
        if (!TextUtils.isEmpty(wifiIp)) {
            tvRetry.setVisibility(View.GONE);
            mTvWifiIp.setText("http://" + NetworkUtils.getConnectWifiIp(WifiBookActivity.this) + ":" + Defaults.getPort());
            // 启动wifi传书服务器
            ServerRunner.startServer(this);
        } else {
            mTvWifiIp.setText("请开启Wifi并重试");
            tvRetry.setVisibility(View.VISIBLE);
        }
        tvRetry.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                initDatas();
            }
        });
    }

    @Override
    public void onBackPressed() {
        if (ServerRunner.serverIsRunning) {
            new AlertDialog.Builder(this)
                    .setTitle("提示")
                    .setMessage("确定要关闭？Wifi传书将会中断！")
                    .setPositiveButton("确定", new DialogInterface.OnClickListener() {
                        @Override
                        public void onClick(DialogInterface dialog, int which) {
                            dialog.dismiss();
                            finish();
                        }
                    }).setNegativeButton("取消", new DialogInterface.OnClickListener() {
                @Override
                public void onClick(DialogInterface dialog, int which) {
                    dialog.dismiss();
                }
            }).create().show();
        } else {
            super.onBackPressed();
        }
    }

    @Override
    protected void onDestroy() {
        if (null != alertDialog) {
            alertDialog.dismiss();
            alertDialog = null;
        }
        super.onDestroy();
        ServerRunner.stopServer();
    }

    AlertDialog alertDialog;

    @Override
    public void SendSuccess(String file) {
        if (null != alertDialog) alertDialog.dismiss();
        alertDialog = new AlertDialog.Builder(this)
                .setTitle(file)
                .setMessage("上传成功,点确定查看")
                .setPositiveButton("确定", new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface dialog, int which) {
                        dialog.dismiss();
                        Intent intent = new Intent(WifiBookActivity.this, ReadFActivity.class);
                        intent.setFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP | Intent.FLAG_ACTIVITY_SINGLE_TOP);
                        startActivity(intent);
                        finish();
                    }
                }).setNegativeButton("取消", new DialogInterface.OnClickListener() {
                    @Override
                    public void onClick(DialogInterface dialog, int which) {
                        dialog.dismiss();
                    }
                }).create();
        alertDialog.show();
        new Handler(Looper.getMainLooper()).postDelayed(new Runnable() {
            @Override
            public void run() {
                alertDialog.dismiss();
            }
        }, 30 * 1000);
    }


    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        getMenuInflater().inflate(R.menu.menu, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        switch (item.getItemId()) {
            case R.id.share:
                Intent sendIntent = new Intent();
                sendIntent.setAction(Intent.ACTION_SEND);
                sendIntent.putExtra(Intent.EXTRA_TEXT, "" + mTvWifiIp.getText());
                sendIntent.setType("text/plain");
                startActivity(sendIntent);
                break;
            case R.id.back:
                finish();
                break;
        }
        return true;
    }
}
