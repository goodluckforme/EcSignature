package com.example.readpoi;

import java.io.*;
import java.util.List;
import java.util.zip.*;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;



/*import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;*/

import android.os.Bundle;
import android.app.Activity;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.util.Log;
import android.util.Xml;
import android.view.Menu;
import android.webkit.WebView;
import android.widget.TextView;

public class ReadWord extends Activity {
	private TextView tv;
	private String nameStr = null;
	private Range range = null;
	private HWPFDocument hwpf = null;
	private String htmlPath;
	private String picturePath;
	private WebView view;
	private List pictures;
	private TableIterator tableIterator;
	private int presentPicture = 0;
	private int screenWidth;
	private FileOutputStream output;
	private File myFile;
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_read_word);
		tv=(TextView) findViewById(R.id.TV2);
		/*this.webView = (WebView)findViewById(R.id.WebView);
		this.webView.getSettings().setBuiltInZoomControls(true);
		this.webView.getSettings().setUseWideViewPort(true);
		this.webView.getSettings().setSupportZoom(true);*/
		//	readWord2007();
		//String x=ReadWord.readDOCX();
		//tv.setText(x);

		this.makeFile();

		this.docx();
		tv.setText(this.htmlPath);
		this.view.loadUrl(this.htmlPath);

		System.out.println("htmlPath" + this.htmlPath);

	}

	/*用来在sdcard上创建图片*/
	public void makePictureFile(){
		String sdString = android.os.Environment.getExternalStorageState();//获取外部存储状态
		//if(sdString.equals(android.os.Environment.MEDIA_MOUNTED)){//确认sd卡存在,原理不知
		try{
			//		File picFile = android.os.Environment.getExternalStorageDirectory();//获取sd卡目录
			//File picFile = new File("D:\\");
			//	String picPath = picFile.getAbsolutePath() + File.separator + "xiao";//创建目录,上面有解释
			File picDirFile = new File("D:/xiao");
			if(!picDirFile.exists()){
				picDirFile.mkdir();
			}
			File pictureFile = new File("D:\\xiao\\" + presentPicture + ".jpg");//创建jpg文件,方法与html相同
			if(!pictureFile.exists()){
				pictureFile.createNewFile();
			}
			this.picturePath = pictureFile.getAbsolutePath();//获取jpg文件绝对路径
		}
		catch(Exception e){
			System.out.println("PictureFile Catch Exception");
			//	}
		}
	}


	public void makeFile(){
		//String sdStateString = android.os.Environment.getExternalStorageState();//获取外部存储状态
//if(sdStateString.equals(android.os.Environment.MEDIA_MOUNTED)){//确认sd卡存在,原理不知,媒体安装??
		try{
			//File sdFile = android.os.Environment.getExternalStorageDirectory();//获取扩展设备的文件目录
			//File sdFile = new File("D:\\");//获取扩展设备的文件目录
			//String path = sdFile.getAbsolutePath() + File.separator + "xiao";//得到sd卡(扩展设备)的绝对路径+"/"+xiao
			//File dirFile = new File(path);//获取xiao文件夹地址
			File dirFile = new File("D:\\A");
			if(!dirFile.exists()){//如果不存在
				dirFile.mkdir();//创建目录
			}
			File myFile = new File("D:\\A\\my.html");//获取my.html的地址
			if(!myFile.exists()){//如果不存在
				myFile.createNewFile();//创建文件
			}
			this.htmlPath = myFile.getAbsolutePath();//返回路径
		}
		catch(Exception e){
		}
		//}
	}
	public void writePicture(){
		Picture picture = (Picture)this.pictures.get(this.presentPicture);
		byte[] pictureBytes = picture.getContent();
		Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0, pictureBytes.length);
		makePictureFile();
		this.presentPicture++;
		File myPicture = new File(this.picturePath);
		try{
			FileOutputStream outputPicture = new FileOutputStream(myPicture);
			outputPicture.write(pictureBytes);
			outputPicture.close();
		}
		catch(Exception e){
			System.out.println("outputPicture Exception");
		}
		String imageString = "<img src=\"" + picturePath + "\"";
		if(bitmap.getWidth() > this.screenWidth){
			imageString = imageString + " " + "width=\"" + this.screenWidth + "\"";
		}
		imageString = imageString + ">";
		try{
			this.output.write(imageString.getBytes());
		}
		catch(Exception e){
			System.out.println("output Exception");
		}
	}

	/*public static String readDOCX() {
		 String path=android.os.Environment.getExternalStorageDirectory()+"/docx.docx";
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

						river += xmlParser.nextText() + "";


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
*/
/*private void readWord2007() {
	 String path=android.os.Environment.getExternalStorageDirectory()+"/docx.docx";

         try {


            OPCPackage oPCPackage = POIXMLDocument.openPackage(path);

            XWPFDocument xwpf = new XWPFDocument(oPCPackage);


            POIXMLTextExtractor ex = new XWPFWordExtractor(xwpf);


           tv.setText(ex.getText()) ;


        } catch (FileNotFoundException e) {


             e.printStackTrace();


         } catch (IOException e) {


            e.printStackTrace();


        }


   }
*/

	public  void docx() {
		String path="D:\\docx1.docx";
		String river = "";
		try {
			this.myFile = new File(this.htmlPath);//new一个File,路径为html文件
			this.output = new FileOutputStream(this.myFile);//new一个流,目标为html文件
			String head = "<!DOCTYPE><html><meta charset=\"utf-8\"><body>";//定义头文件,我在这里加了utf-8,不然会出现乱码
			String end = "</body></html>";
			String tagBegin = "<p>";//段落开始,标记开始?
			String tagEnd = "</p>";//段落结束
			String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
			String tableEnd = "</table>";
			String rowBegin = "<tr>";
			String rowEnd = "</tr>";
			String colBegin = "<td>";
			String colEnd = "</td>";
			this.output.write(head.getBytes());//写如头部


			ZipFile xlsxFile = new ZipFile(new File(path));
			ZipEntry sharedStringXML = xlsxFile.getEntry("word/document.xml");
			InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
			XmlPullParser xmlParser = Xml.newPullParser();
			xmlParser.setInput(inputStream, "utf-8");
			int evtType = xmlParser.getEventType();
			boolean isTable = false; //是表格  用来统计 列 行 数
			boolean isSize = false;  //大小状态
			boolean isColor = false;  //颜色状态
			boolean isCenter = false; //居中状态
			boolean isRight = false; //居右状态
			boolean isItalic = false; //是斜体
			boolean isUnderline = false;  //是下划线
			boolean isBold = false;  //加粗
			boolean isR = false;    //在那个r中
			int pictureIndex = 1;  //docx 压缩包中的图片名  iamge1 开始  所以索引从1开始
			while (evtType != XmlPullParser.END_DOCUMENT) {
				switch (evtType) {

					//开始标签
					case XmlPullParser.START_TAG:
						String tag = xmlParser.getName();
						System.out.println(tag);

						if(tag.equalsIgnoreCase("r")){
							isR = true;
						}
						if(tag.equalsIgnoreCase("u")){  //判断下划线
							isUnderline = true;
						}
						if(tag.equalsIgnoreCase("jc")){  //判断对齐方式
							String align = xmlParser.getAttributeValue(0);
							if(align.equals("center")){
								this.output.write("<center>".getBytes());
								isCenter = true;
							}
							if(align.equals("right")){
								this.output.write("<div align=\"right\">".getBytes());
								isRight = true;
							}
						}
						if(tag.equalsIgnoreCase("color")){   //判断颜色
							String color = xmlParser.getAttributeValue(0);
							this.output.write(("<font color=" + color + ">").getBytes());
							isColor = true;
						}
						if(tag.equalsIgnoreCase("sz")){   //判断大小
							if(isR == true){
								int size = decideSize(Integer.valueOf(xmlParser.getAttributeValue(0)));
								this.output.write(("<font size=" + size + ">").getBytes());
								isSize = true;
							}
						}
						//下面是表格处理
						if(tag.equalsIgnoreCase("tbl")){  //检测到tbl  表格开始
							this.output.write(tableBegin.getBytes());
							isTable = true;
						}
						if(tag.equalsIgnoreCase("tr")){    //行
							this.output.write(rowBegin.getBytes());
						}
						if(tag.equalsIgnoreCase("tc")){   //列
							this.output.write(colBegin.getBytes());
						}

						if(tag.equalsIgnoreCase("pict")){  //检测到标签  pict  图片
							ZipEntry sharePicture = xlsxFile.getEntry("word/media/image" + pictureIndex + ".jpeg");//一下为读取docx的图片  转化为流数组
							InputStream pictIS = xlsxFile.getInputStream(sharePicture);
							ByteArrayOutputStream pOut = new ByteArrayOutputStream();
							byte [] bt = null;
							byte [] b = new byte[1000];
							int len = 0;
							while ((len = pictIS.read(b)) != -1) {
								pOut.write(b, 0, len);
							}
							pictIS.close();
							pOut.close();
							bt = pOut.toByteArray();
							Log.i("byteArray", ""+bt);
							if (pictIS != null)
								pictIS.close();
							if (pOut != null)
								pOut.close();
							writeDOCXPicture(bt);

							pictureIndex ++;  //转换一张后 索引+1
						}

						if(tag.equalsIgnoreCase("b")){  //检测到加粗标签
							isBold = true;
						}
						if(tag.equalsIgnoreCase("p")){//检测到 p 标签
							if(isTable == false){     // 如果在表格中 就无视
								this.output.write(tagBegin.getBytes());
							}
						}
						if(tag.equalsIgnoreCase("i")){   //斜体
							isItalic = true;
						}
						//检测到值 标签
						if (tag.equalsIgnoreCase("t")) {
							if(isBold == true){   //加粗
								this.output.write("<b>".getBytes());
							}
							if(isUnderline == true){    //检测到下划线标签,输入<u>
								this.output.write("<u>".getBytes());
							}
							if(isItalic == true){       //检测到斜体标签,输入<i>
								output.write("<i>".getBytes());
							}
							river = xmlParser.nextText();
							this.output.write(river.getBytes());  //写入数值
							if(isItalic == true){      //检测到斜体标签,在输入值之后,输入</i>,并且斜体状态=false
								this.output.write("</i>".getBytes());
								isItalic = false;
							}
							if(isUnderline == true){//检测到下划线标签,在输入值之后,输入</u>,并且下划线状态=false
								this.output.write("</u>".getBytes());
								isUnderline = false;
							}
							if(isBold == true){   //加粗
								this.output.write("</b>".getBytes());
								isBold = false;
							}
							if(isSize == true){   //检测到大小设置,输入结束标签
								this.output.write("</font>".getBytes());
								isSize = false;
							}
							if(isColor == true){  //检测到颜色设置存在,输入结束标签
								this.output.write("</font>".getBytes());
								isColor = false;
							}
							if(isCenter == true){   //检测到居中,输入结束标签
								this.output.write("</center>".getBytes());
								isCenter = false;
							}
							if(isRight == true){  //居右不能使用<right></right>,使用div可能会有状况,先用着
								this.output.write("</div>".getBytes());
								isRight = false;
							}
						}
						break;
					//结束标签
					case XmlPullParser.END_TAG:
						String tag2 = xmlParser.getName();
						if(tag2.equalsIgnoreCase("tbl")){  //检测到表格结束,更改表格状态
							this.output.write(tableEnd.getBytes());
							isTable = false;
						}
						if(tag2.equalsIgnoreCase("tr")){  //行结束
							this.output.write(rowEnd.getBytes());
						}
						if(tag2.equalsIgnoreCase("tc")){  //列结束
							this.output.write(colEnd.getBytes());
						}
						if(tag2.equalsIgnoreCase("p")){   //p结束,如果在表格中就无视
							if(isTable == false){
								this.output.write(tagEnd.getBytes());
							}
						}
						if(tag2.equalsIgnoreCase("r")){
							isR = false;
						}
						break;
					default:
						break;
				}
				evtType = xmlParser.next();
			}
			this.output.write(end.getBytes());
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
	}
	public int decideSize(int size){
		if(size >= 1 && size <= 8){
			return 1;
		}
		if(size >= 9 && size <= 11){
			return 2;
		}
		if(size >= 12 && size <= 14){
			return 3;
		}
		if(size >= 15 && size <= 19){
			return 4;
		}
		if(size >= 20 && size <= 29){
			return 5;
		}
		if(size >= 30 && size <= 39){
			return 6;
		}
		if(size >= 40){
			return 7;
		}
		return 3;
	}
	public void writeDOCXPicture(byte[] pictureBytes){
		Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0, pictureBytes.length);
		makePictureFile();
		presentPicture++;
		File myPicture = new File(this.picturePath);
		try{
			FileOutputStream outputPicture = new FileOutputStream(myPicture);
			outputPicture.write(pictureBytes);
			outputPicture.close();
		}
		catch(Exception e){
			System.out.println("outputPicture Exception");
		}
		String imageString = "<img src=\"" + this.picturePath + "\"";
		if(bitmap.getWidth() > this.screenWidth){
			imageString = imageString + " " + "width=\"" + this.screenWidth + "\"";
		}
		imageString = imageString + ">";
		try{
			this.output.write(imageString.getBytes());
		}
		catch(Exception e){
			System.out.println("output Exception");
		}
	}


}
