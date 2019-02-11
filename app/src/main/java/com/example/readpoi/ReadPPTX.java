package com.example.readpoi;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import android.app.Activity;
import android.os.Bundle;
import android.util.Xml;
import android.view.Menu;
import android.widget.ImageView;
import android.widget.TextView;

public class ReadPPTX extends Activity {
	private TextView tv;
	ImageView img;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_read_pptx);
		img = (ImageView) findViewById(R.id.pptximg);
		tv = (TextView) findViewById(R.id.pptxtv);
		tv.setText(readPPTX());
	}



	public String readPPTX() {
		String path = android.os.Environment.getExternalStorageDirectory()
				+ "/cc.pptx";
		List<String> ls = new ArrayList<String>();
		String river = "";
		ZipFile xlsxFile = null;
		try {
			xlsxFile = new ZipFile(new File(path));// pptx按照读取zip格式读取
		} catch (ZipException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		try {
			ZipEntry sharedStringXML = xlsxFile.getEntry("[Content_Types].xml");// 找到里面存放内容的文件
			InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);// 将得到文件流
			XmlPullParser xmlParser = Xml.newPullParser();// 实例化pull
			xmlParser.setInput(inputStream, "utf-8");// 将流放进pull中
			int evtType = xmlParser.getEventType();// 得到标签类型的状态
			while (evtType != XmlPullParser.END_DOCUMENT) {// 循环读取流
				switch (evtType) {
					case XmlPullParser.START_TAG: // 判断标签开始读取
						String tag = xmlParser.getName();// 得到标签
						if (tag.equalsIgnoreCase("Override")) {
							String s = xmlParser
									.getAttributeValue(null, "PartName");
							if (s.lastIndexOf("/ppt/slides/slide") == 0) {
								ls.add(s);
							}
						}
						break;
					case XmlPullParser.END_TAG:// 标签读取结束
						break;
					default:
						break;
				}
				evtType = xmlParser.next();// 读取下一个标签
			}
		} catch (ZipException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (XmlPullParserException e) {
			e.printStackTrace();
		}
		for (int i = 1; i < (ls.size() + 1); i++) {// 假设有6张幻灯片
			river += "第" + i + "张:" + "\n";
			try {
				ZipEntry sharedStringXML = xlsxFile.getEntry("ppt/slides/slide"
						+ i + ".xml");// 找到里面存放内容的文件
				InputStream inputStream = xlsxFile
						.getInputStream(sharedStringXML);// 将得到文件流
				XmlPullParser xmlParser = Xml.newPullParser();// 实例化pull
				xmlParser.setInput(inputStream, "utf-8");// 将流放进pull中
				int evtType = xmlParser.getEventType();// 得到标签类型的状态
				while (evtType != XmlPullParser.END_DOCUMENT) {// 循环读取流
					switch (evtType) {
						case XmlPullParser.START_TAG: // 判断标签开始读取
							String tag = xmlParser.getName();// 得到标签
							if (tag.equalsIgnoreCase("t")) {
								river += xmlParser.nextText() + "\n";

							} else if (tag.equalsIgnoreCase("cNvPr")) {

								// img.setImageResource();
							}
							break;
						case XmlPullParser.END_TAG:// 标签读取结束
							break;
						default:
							break;
					}
					evtType = xmlParser.next();// 读取下一个标签
				}
			} catch (ZipException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (XmlPullParserException e) {
				e.printStackTrace();
			}
		}
		if (river == null) {
			river = "解析文件出现问题";
		}
		return river;
	}
}
