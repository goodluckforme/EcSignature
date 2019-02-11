/**
 * Copyright 2016 JustWayward Team
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.example.wifitransfer;
import com.example.readpoi.WifiBookActivity;

import java.io.IOException;

/**
 * Wifi传书 服务端
 *
 * @author yuyh.
 * @date 2016/10/10.
 */
public class ServerRunner {

    private static SimpleFileServer server;
    public static boolean serverIsRunning = false;

    /**
     * 启动wifi传书服务
     * @param sendListener
     */
    public static void startServer(FileSendListener sendListener) {
        server = SimpleFileServer.getInstance();
        try {
            if (!serverIsRunning) {
                server.start(sendListener);
                serverIsRunning = true;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void stopServer() {
        if (server != null) {
            server.stop();
            serverIsRunning = false;
        }
    }
}