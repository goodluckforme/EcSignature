package com.example.readpoi;

import android.app.Application;
import android.content.Context;
import android.support.multidex.MultiDex;

/**
 * Created by Administrator on 2019/1/10.
 */

public class App extends Application {

    @Override
    protected void attachBaseContext(Context base) {
        super.attachBaseContext(base);
        MultiDex.install(this);
    }

    public static Application inistance;

    @Override
    public void onCreate() {
        super.onCreate();
        inistance = this;
    }
}
