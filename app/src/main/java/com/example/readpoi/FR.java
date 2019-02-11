package com.example.readpoi;

import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.os.Environment;
import android.util.Log;
import android.util.Xml;

import com.example.signature.RxFileTool;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

public class FR {
    private String fileNameNoExtension;
    private String nameStr;
    public Range range = null;
    public HWPFDocument hwpf = null;
    public String htmlPath;
    public String picturePath;
    public List pictures;
    public TableIterator tableIterator;
    public int presentPicture = 0;
    public int screenWidth;
    public FileOutputStream output;
    public File myFile;
    StringBuffer lsb = new StringBuffer();
    String returnPath = "";
    static final int BUFFER = 2048;
    String picPath = "";// 创建目录,上面有解释

    public FR(String namepath) {
        // this.screenWidth =
        // this.getWindowManager().getDefaultDisplay().getWidth() -
        // 10;//设置宽度为屏幕宽度-10
        this.nameStr = namepath;
        fileNameNoExtension = RxFileTool.getFileNameNoExtension(nameStr);
        picPath = Environment
                .getExternalStorageDirectory().getAbsolutePath() + File.separator
                + "piopic" + File.separator
                + fileNameNoExtension;
        read();
    }

    public void read() {

        if (this.nameStr.endsWith(".doc")) {
            this.getRange();
            this.makeFile();
            this.readDOC();
            returnPath = "file:///" + this.htmlPath;
            // this.webView.loadUrl("file:///" + this.htmlPath);
            System.out.println("htmlPath" + this.htmlPath);
        }
        if (this.nameStr.endsWith(".docx")) {
            this.makeFile();
            this.readDOCX();
            returnPath = "file:///" + this.htmlPath;
            // this.webView.loadUrl("file:///" + this.htmlPath);
            System.out.println("htmlPath" + this.htmlPath);
        }
        if (this.nameStr.endsWith(".xls")) {

            try {
                this.makeFile();
                this.readXLS();
                returnPath = "file:///" + this.htmlPath;
                // this.webView.loadUrl("file:///" + this.htmlPath);
                System.out.println("htmlPath" + this.htmlPath);
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
        if (this.nameStr.endsWith(".xlsx")) {
            this.makeFile();
            this.readXLSX();
            returnPath = "file:///" + this.htmlPath;
            // this.webView.loadUrl("file:///" + this.htmlPath);
            System.out.println("htmlPath" + this.htmlPath);
        }

    }

    /* 读取word中的内容写到sdcard上的.html文件中 */
    public void readDOC() {

        try {
            myFile = new File(htmlPath);
            output = new FileOutputStream(myFile);
//            String head = "<html><meta charset=\"utf-8\"><body>";
            String head = "<!DOCTYPE>\n" +
                    "<html>\n" +
                    "\n" +
                    "\t<head>\n" +
                    "\t\t<meta charset=\"utf-8\">\n" +
                    "\t\t<meta name=\"viewport\" content=\"width=device-width, initial-scale=1,maximum-scale=1,user-scalable=yes\">\n" +
                    "\t\t<style type=\"text/css\">\n" +
                    "\t\t\t\timg {\n" +
                    "\t\t\t\twidth: 100%;\n" +
                    "\t\t\t\theight: auto;\n" +
                    "\t\t\t\tvertical-align: middle;\n" +
                    "\t\t</style>\n" +
                    "\t</head>\n" +
                    "\n" +
                    "\t<body>";
            String tagBegin = "<p>";
            String tagEnd = "</p>";
            output.write(head.getBytes());
            int numParagraphs = range.numParagraphs();// 得到页面所有的段落数
            for (int i = 0; i < numParagraphs; i++) { // 遍历段落数
                Paragraph p = range.getParagraph(i); // 得到文档中的每一个段落
                if (p.isInTable()) {
                    int temp = i;
                    if (tableIterator.hasNext()) {
                        String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
                        String tableEnd = "</table>";
                        String rowBegin = "<tr>";
                        String rowEnd = "</tr>";
                        String colBegin = "<td>";
                        String colEnd = "</td>";
                        Table table = tableIterator.next();
                        output.write(tableBegin.getBytes());
                        int rows = table.numRows();
                        for (int r = 0; r < rows; r++) {
                            output.write(rowBegin.getBytes());
                            TableRow row = table.getRow(r);
                            int cols = row.numCells();
                            int rowNumParagraphs = row.numParagraphs();
                            int colsNumParagraphs = 0;
                            for (int c = 0; c < cols; c++) {
                                output.write(colBegin.getBytes());
                                TableCell cell = row.getCell(c);
                                int max = temp + cell.numParagraphs();
                                colsNumParagraphs = colsNumParagraphs
                                        + cell.numParagraphs();
                                for (int cp = temp; cp < max; cp++) {
                                    Paragraph p1 = range.getParagraph(cp);
                                    output.write(tagBegin.getBytes());
                                    writeParagraphContent(p1);
                                    output.write(tagEnd.getBytes());
                                    temp++;
                                }
                                output.write(colEnd.getBytes());
                            }
                            int max1 = temp + rowNumParagraphs;
                            for (int m = temp + colsNumParagraphs; m < max1; m++) {
                                temp++;
                            }
                            output.write(rowEnd.getBytes());
                        }
                        output.write(tableEnd.getBytes());
                    }
                    i = temp;
                } else {
                    output.write(tagBegin.getBytes());
                    writeParagraphContent(p);
                    output.write(tagEnd.getBytes());
                }
            }
            String end = "<p>\n" +
                    "\t\t\t<center><img id=\"signature\" src=\"" +
                    "../signature.png\"" +
                    "></p>\n" +
                    "\t\t<p>\n" +
                    "\t\t\t<script type=\"text/javascript\">\n" +
                    "\t\t\t\tdocument.getElementById('signature').onclick = function() {\n" +
                    "\t\t\t\t\tcallAndroid.showPad();\n" +
                    "\t\t\t\t\tconsole.log(\"showPad\");\n" +
                    "\t\t\t\t}\n" +
                    "\n" +
                    "\t\t\t\tfunction setImgToWb(path) {\n" +
                    "\t\t\t\t\tconsole.log(path);\n" +
                    "\t\t\t\t\tdocument.getElementById('signature').src = path\n" +
                    "\t\t\t\t}\n" +
                    "\t\t\t</script>"
                    + "</body></html>";
            output.write(end.getBytes());
            output.close();
        } catch (Exception e) {

            System.out.println("readAndWrite Exception:" + e.getMessage());
            e.printStackTrace();
        }
    }

    public void readDOCX() {
        String river = "";
        try {
            this.myFile = new File(this.htmlPath);// new一个File,路径为html文件
            this.output = new FileOutputStream(this.myFile);// new一个流,目标为html文件
//            String head = "<!DOCTYPE><html><meta charset=\"utf-8\"><body>";// 定义头文件,我在这里加了utf-8,不然会出现乱码
            String head = "<!DOCTYPE>\n" +
                    "<html>\n" +
                    "\n" +
                    "\t<head>\n" +
                    "\t\t<meta charset=\"utf-8\">\n" +
                    "\t\t<meta name=\"viewport\" content=\"width=device-width, initial-scale=1,maximum-scale=1,user-scalable=yes\">\n" +
                    "\t\t<style type=\"text/css\">\n" +
                    "\t\t\t\timg {\n" +
                    "\t\t\t\twidth: 100%;\n" +
                    "\t\t\t\theight: auto;\n" +
                    "\t\t\t\tvertical-align: middle;\n" +
                    "\t\t</style>\n" +
                    "\t</head>\n" +
                    "\n" +
                    "\t<body>";// 定义头文件,我在这里加了utf-8,不然会出现乱码
            String end = "<p>\n" +
                    "\t\t\t<center><img id=\"signature\" src=\"" +
                    "../signature.png" +
                    "\"></p>\n" +
                    "<script type=\"text/javascript\">\n" +
                    "\t\t\t\tdocument.getElementById(\"signature\").onclick = function() {\n" +
                    "\t\t\t\t\tcallAndroid.showPad();\n" +
                    "\t\t\t\t\tconsole.log(\"showPad\");\n" +
                    "\t\t\t\t}\n" +
                    "\n" +
                    "\t\t\t\tfunction setImgToWb(path) {\n" +
                    "\t\t\t\t\tconsole.log(path);\n" +
                    "\t\t\t\t\tdocument.getElementById(\"signature\").src = path\n" +
                    "\t\t\t\t}\n" +
                    "\t\t\t</script>" +
                    "\t</body></html>";
            String tagBegin = "<p>";// 段落开始,标记开始?
            String tagEnd = "</p>";// 段落结束
            String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
            String tableEnd = "</table>";
            String rowBegin = "<tr>";
            String rowEnd = "</tr>";
            String colBegin = "<td>";
            String colEnd = "</td>";
            String style = "style=\"";
            this.output.write(head.getBytes());// 写如头部
            ZipFile xlsxFile = new ZipFile(new File(this.nameStr));
            ZipEntry sharedStringXML = xlsxFile.getEntry("word/document.xml");
            InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
            XmlPullParser xmlParser = Xml.newPullParser();
            xmlParser.setInput(inputStream, "utf-8");
            int evtType = xmlParser.getEventType();
            boolean isTable = false; // 是表格 用来统计 列 行 数
            boolean isSize = false; // 大小状态
            boolean isColor = false; // 颜色状态
            boolean isCenter = false; // 居中状态
            boolean isRight = false; // 居右状态
            boolean isItalic = false; // 是斜体
            boolean isUnderline = false; // 是下划线
            boolean isBold = false; // 加粗
            boolean isR = false; // 在那个r中
            boolean isStyle = false;
            int pictureIndex = 1; // docx 压缩包中的图片名 iamge1 开始 所以索引从1开始
            while (evtType != XmlPullParser.END_DOCUMENT) {
                switch (evtType) {

                    // 开始标签
                    case XmlPullParser.START_TAG:
                        String tag = xmlParser.getName();

                        if (tag.equalsIgnoreCase("r")) {
                            isR = true;
                        }
                        if (tag.equalsIgnoreCase("u")) { // 判断下划线
                            isUnderline = true;
                        }
                        if (tag.equalsIgnoreCase("jc")) { // 判断对齐方式
                            String align = xmlParser.getAttributeValue(0);
                            if (align.equals("center")) {
                                this.output.write("<center>".getBytes());
                                isCenter = true;
                            }
                            if (align.equals("right")) {
                                this.output.write("<div align=\"right\">"
                                        .getBytes());
                                isRight = true;
                            }
                        }

                        if (tag.equalsIgnoreCase("color")) { // 判断颜色

                            String color = xmlParser.getAttributeValue(0);

                            this.output
                                    .write(("<span style=\"color:" + color + ";\">")
                                            .getBytes());
                            isColor = true;
                        }
                        if (tag.equalsIgnoreCase("sz")) { // 判断大小
                            if (isR == true) {
                                int size = decideSize(Integer.valueOf(xmlParser
                                        .getAttributeValue(0)));
                                this.output.write(("<font size=" + size + ">")
                                        .getBytes());
                                isSize = true;
                            }
                        }
                        // 下面是表格处理
                        if (tag.equalsIgnoreCase("tbl")) { // 检测到tbl 表格开始
                            this.output.write(tableBegin.getBytes());
                            isTable = true;
                        }
                        if (tag.equalsIgnoreCase("tr")) { // 行
                            this.output.write(rowBegin.getBytes());
                        }
                        if (tag.equalsIgnoreCase("tc")) { // 列
                            this.output.write(colBegin.getBytes());
                        }

                        if (tag.equalsIgnoreCase("pic")) { // 检测到标签 pic 图片
                            String entryName_jpeg = "word/media/image"
                                    + pictureIndex + ".jpeg";
                            String entryName_png = "word/media/image"
                                    + pictureIndex + ".png";
                            String entryName_gif = "word/media/image"
                                    + pictureIndex + ".gif";
                            String entryName_wmf = "word/media/image"
                                    + pictureIndex + ".wmf";
                            ZipEntry sharePicture = null;
                            InputStream pictIS = null;
                            sharePicture = xlsxFile.getEntry(entryName_jpeg);
                            // 一下为读取docx的图片 转化为流数组
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_png);
                            }
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_gif);
                            }
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_wmf);
                            }

                            if (sharePicture != null) {
                                pictIS = xlsxFile.getInputStream(sharePicture);
                                ByteArrayOutputStream pOut = new ByteArrayOutputStream();
                                byte[] bt = null;
                                byte[] b = new byte[1000];
                                int len = 0;
                                while ((len = pictIS.read(b)) != -1) {
                                    pOut.write(b, 0, len);
                                }
                                pictIS.close();
                                pOut.close();
                                bt = pOut.toByteArray();
                                Log.i("byteArray", "" + bt);
                                if (pictIS != null)
                                    pictIS.close();
                                if (pOut != null)
                                    pOut.close();
                                writeDOCXPicture(bt);
                            }

                            pictureIndex++; // 转换一张后 索引+1
                        }

                        if (tag.equalsIgnoreCase("b")) { // 检测到加粗标签
                            isBold = true;
                        }
                        if (tag.equalsIgnoreCase("p")) {// 检测到 p 标签
                            if (isTable == false) { // 如果在表格中 就无视
                                this.output.write(tagBegin.getBytes());
                            }
                        }
                        if (tag.equalsIgnoreCase("i")) { // 斜体
                            isItalic = true;
                        }
                        // 检测到值 标签
                        if (tag.equalsIgnoreCase("t")) {
                            if (isBold == true) { // 加粗
                                this.output.write("<b>".getBytes());
                            }
                            if (isUnderline == true) { // 检测到下划线标签,输入<u>
                                this.output.write("<u>".getBytes());
                            }
                            if (isItalic == true) { // 检测到斜体标签,输入<i>
                                output.write("<i>".getBytes());
                            }
                            river = xmlParser.nextText();
                            this.output.write(river.getBytes()); // 写入数值
                            if (isItalic == true) { // 检测到斜体标签,在输入值之后,输入</i>,并且斜体状态=false
                                this.output.write("</i>".getBytes());
                                isItalic = false;
                            }
                            if (isUnderline == true) {// 检测到下划线标签,在输入值之后,输入</u>,并且下划线状态=false
                                this.output.write("</u>".getBytes());
                                isUnderline = false;
                            }
                            if (isBold == true) { // 加粗
                                this.output.write("</b>".getBytes());
                                isBold = false;
                            }
                            if (isSize == true) { // 检测到大小设置,输入结束标签
                                this.output.write("</font>".getBytes());
                                isSize = false;
                            }
                            if (isColor == true) { // 检测到颜色设置存在,输入结束标签
                                this.output.write("</span>".getBytes());
                                isColor = false;
                            }
                            if (isCenter == true) { // 检测到居中,输入结束标签
                                this.output.write("</center>".getBytes());
                                isCenter = false;
                            }
                            if (isRight == true) { // 居右不能使用<right></right>,使用div可能会有状况,先用着
                                this.output.write("</div>".getBytes());
                                isRight = false;
                            }
                        }
                        break;
                    // 结束标签
                    case XmlPullParser.END_TAG:
                        String tag2 = xmlParser.getName();
                        if (tag2.equalsIgnoreCase("tbl")) { // 检测到表格结束,更改表格状态
                            this.output.write(tableEnd.getBytes());
                            isTable = false;
                        }
                        if (tag2.equalsIgnoreCase("tr")) { // 行结束
                            this.output.write(rowEnd.getBytes());
                        }
                        if (tag2.equalsIgnoreCase("tc")) { // 列结束
                            this.output.write(colEnd.getBytes());
                        }
                        if (tag2.equalsIgnoreCase("p")) { // p结束,如果在表格中就无视
                            if (isTable == false) {
                                this.output.write(tagEnd.getBytes());
                            }
                        }
                        if (tag2.equalsIgnoreCase("r")) {
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

    public StringBuffer readXLS() throws Exception {

        myFile = new File(htmlPath);
        output = new FileOutputStream(myFile);
        lsb.append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>");
        lsb.append("<head><meta http-equiv=Content-Type content='text/html; charset=utf-8'><meta name=ProgId content=Excel.Sheet>");
        HSSFSheet sheet = null;

        String excelFileName = nameStr;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(
                    excelFileName)); // 获整个Excel

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                sheet = workbook.getSheetAt(sheetIndex);// 获所有的sheet
                String sheetName = workbook.getSheetName(sheetIndex); // sheetName
                if (workbook.getSheetAt(sheetIndex) != null) {
                    sheet = workbook.getSheetAt(sheetIndex);// 获得不为空的这个sheet
                    if (sheet != null) {
                        int firstRowNum = sheet.getFirstRowNum(); // 第一行
                        int lastRowNum = sheet.getLastRowNum(); // 最后一行
                        // 构造Table
                        lsb.append("<table width=\"100%\" style=\"border:1px solid #000;border-width:1px 0 0 1px;margin:2px 0 2px 0;border-collapse:collapse;\">");
                        for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
                            if (sheet.getRow(rowNum) != null) {// 如果行不为空，
                                HSSFRow row = sheet.getRow(rowNum);
                                short firstCellNum = row.getFirstCellNum(); // 该行的第一个单元格
                                short lastCellNum = row.getLastCellNum(); // 该行的最后一个单元格
                                int height = (int) (row.getHeight() / 15.625); // 行的高度
                                lsb.append("<tr height=\""
                                        + height
                                        + "\" style=\"border:1px solid #000;border-width:0 1px 1px 0;margin:2px 0 2px 0;\">");
                                for (short cellNum = firstCellNum; cellNum <= lastCellNum; cellNum++) { // 循环该行的每一个单元格
                                    HSSFCell cell = row.getCell(cellNum);
                                    if (cell != null) {
                                        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                                            continue;
                                        } else {
                                            StringBuffer tdStyle = new StringBuffer(
                                                    "<td style=\"border:1px solid #000; border-width:0 1px 1px 0;margin:2px 0 2px 0; ");
                                            HSSFCellStyle cellStyle = cell
                                                    .getCellStyle();
                                            HSSFPalette palette = workbook
                                                    .getCustomPalette(); // 类HSSFPalette用于求颜色的国际标准形式
                                            HSSFColor hColor = palette
                                                    .getColor(cellStyle
                                                            .getFillForegroundColor());
                                            HSSFColor hColor2 = palette
                                                    .getColor(cellStyle
                                                            .getFont(workbook)
                                                            .getColor());

                                            String bgColor = convertToStardColor(hColor);// 背景颜色
                                            short boldWeight = cellStyle
                                                    .getFont(workbook)
                                                    .getBoldweight(); // 字体粗细
                                            short fontHeight = (short) (cellStyle
                                                    .getFont(workbook)
                                                    .getFontHeight() / 2); // 字体大小
                                            String fontColor = convertToStardColor(hColor2); // 字体颜色
                                            if (bgColor != null
                                                    && !"".equals(bgColor
                                                    .trim())) {
                                                tdStyle.append(" background-color:"
                                                        + bgColor + "; ");
                                            }
                                            if (fontColor != null
                                                    && !"".equals(fontColor
                                                    .trim())) {
                                                tdStyle.append(" color:"
                                                        + fontColor + "; ");
                                            }
                                            tdStyle.append(" font-weight:"
                                                    + boldWeight + "; ");
                                            tdStyle.append(" font-size: "
                                                    + fontHeight + "%;");
                                            lsb.append(tdStyle + "\"");

                                            int width = (int) (sheet
                                                    .getColumnWidth(cellNum) / 35.7); //
                                            int cellReginCol = getMergerCellRegionCol(
                                                    sheet, rowNum, cellNum); // 合并的列（solspan）
                                            int cellReginRow = getMergerCellRegionRow(
                                                    sheet, rowNum, cellNum);// 合并的行（rowspan）
                                            String align = convertAlignToHtml(cellStyle
                                                    .getAlignment()); //
                                            String vAlign = convertVerticalAlignToHtml(cellStyle
                                                    .getVerticalAlignment());

                                            lsb.append(" align=\"" + align
                                                    + "\" valign=\"" + vAlign
                                                    + "\" width=\"" + width
                                                    + "\" ");

                                            lsb.append(" colspan=\""
                                                    + cellReginCol
                                                    + "\" rowspan=\""
                                                    + cellReginRow + "\"");
                                            lsb.append(">" + getCellValue(cell)
                                                    + "</td>");
                                        }
                                    }
                                }
                                lsb.append("</tr>");
                            }
                        }
                    }

                }

            }
            output.write(lsb.toString().getBytes());
        } catch (FileNotFoundException e) {
            throw new Exception("文件 " + excelFileName + " 没有找到!");
        } catch (IOException e) {
            throw new Exception("文件 " + excelFileName + " 处理错误("
                    + e.getMessage() + ")!");
        }
        return lsb;
    }

    public void readXLSX() {
        try {
            this.myFile = new File(this.htmlPath);// new一个File,路径为html文件
            this.output = new FileOutputStream(this.myFile);// new一个流,目标为html文件
            String head = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\"\"http://www.w3.org/TR/html4/loose.dtd\"><html><meta charset=\"utf-8\"><head></head><body>";// 定义头文件,我在这里加了utf-8,不然会出现乱码
            String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
            String tableEnd = "</table>";
            String rowBegin = "<tr>";
            String rowEnd = "</tr>";
            String colBegin = "<td>";
            String colEnd = "</td>";
            String end = "</body></html>";
            this.output.write(head.getBytes());
            this.output.write(tableBegin.getBytes());
            String str = "";
            String v = null;
            boolean flat = false;
            List<String> ls = new ArrayList<String>();
            try {
                ZipFile xlsxFile = new ZipFile(new File(this.nameStr));// 地址
                ZipEntry sharedStringXML = xlsxFile
                        .getEntry("xl/sharedStrings.xml");// 共享字符串
                InputStream inputStream = xlsxFile
                        .getInputStream(sharedStringXML);// 输入流 目标上面的共享字符串
                XmlPullParser xmlParser = Xml.newPullParser();// new 解析器
                xmlParser.setInput(inputStream, "utf-8");// 设置解析器类型
                int evtType = xmlParser.getEventType();// 获取解析器的事件类型
                while (evtType != XmlPullParser.END_DOCUMENT) {// 如果不等于 文档结束
                    switch (evtType) {
                        case XmlPullParser.START_TAG: // 标签开始
                            String tag = xmlParser.getName();
                            if (tag.equalsIgnoreCase("t")) {
                                ls.add(xmlParser.nextText());
                            }
                            break;
                        case XmlPullParser.END_TAG: // 标签结束
                            break;
                        default:
                            break;
                    }
                    evtType = xmlParser.next();
                }
                ZipEntry sheetXML = xlsxFile
                        .getEntry("xl/worksheets/sheet1.xml");
                InputStream inputStreamsheet = xlsxFile
                        .getInputStream(sheetXML);
                XmlPullParser xmlParsersheet = Xml.newPullParser();
                xmlParsersheet.setInput(inputStreamsheet, "utf-8");
                int evtTypesheet = xmlParsersheet.getEventType();
                this.output.write(rowBegin.getBytes());
                int i = -1;
                while (evtTypesheet != XmlPullParser.END_DOCUMENT) {
                    switch (evtTypesheet) {
                        case XmlPullParser.START_TAG: // 标签开始
                            String tag = xmlParsersheet.getName();
                            if (tag.equalsIgnoreCase("row")) {
                            } else {
                                if (tag.equalsIgnoreCase("c")) {
                                    String t = xmlParsersheet.getAttributeValue(
                                            null, "t");
                                    if (t != null) {
                                        flat = true;
                                        System.out.println(flat + "有");
                                    } else {// 没有数据时 找了我n年,终于找到了 输入<td></td> 表示空格
                                        this.output.write(colBegin.getBytes());
                                        this.output.write(colEnd.getBytes());
                                        System.out.println(flat + "没有");
                                        flat = false;
                                    }
                                } else {
                                    if (tag.equalsIgnoreCase("v")) {
                                        v = xmlParsersheet.nextText();
                                        this.output.write(colBegin.getBytes());
                                        if (v != null) {
                                            if (flat) {
                                                str = ls.get(Integer.parseInt(v));
                                            } else {
                                                str = v;
                                            }
                                            this.output.write(str.getBytes());
                                            this.output.write(colEnd.getBytes());
                                        }
                                    }
                                }
                            }
                            break;
                        case XmlPullParser.END_TAG:
                            if (xmlParsersheet.getName().equalsIgnoreCase("row")
                                    && v != null) {
                                if (i == 1) {
                                    this.output.write(rowEnd.getBytes());
                                    this.output.write(rowBegin.getBytes());
                                    i = 1;
                                } else {
                                    this.output.write(rowBegin.getBytes());
                                }
                            }
                            break;
                    }
                    evtTypesheet = xmlParsersheet.next();
                }
                System.out.println(str);
            } catch (ZipException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (XmlPullParserException e) {
                e.printStackTrace();
            }
            if (str == null) {
                str = "解析文件出现问题";
            }
            this.output.write(rowEnd.getBytes());
            this.output.write(tableEnd.getBytes());
            this.output.write(end.getBytes());
        } catch (Exception e) {
            System.out.println("readAndWrite Exception");
        }
    }

    /**
     * 取得单元格的值
     *
     * @param cell
     * @return
     * @throws IOException
     */
    private static Object getCellValue(HSSFCell cell) throws IOException {
        Object value = "";
        if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
            value = cell.getRichStringCellValue().toString();
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                Date date = (Date) cell.getDateCellValue();
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                value = sdf.format(date);
            } else {
                double value_temp = (double) cell.getNumericCellValue();
                BigDecimal bd = new BigDecimal(value_temp);
                BigDecimal bd1 = bd.setScale(3, bd.ROUND_HALF_UP);
                value = bd1.doubleValue();

                DecimalFormat format = new DecimalFormat("#0.###");
                value = format.format(cell.getNumericCellValue());

            }
        }
        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
            value = "";
        }
        return value;
    }

    /**
     * 判断单元格在不在合并单元格范围内，如果是，获取其合并的列数。
     *
     * @param sheet   工作表
     * @param cellRow 被判断的单元格的行号
     * @param cellCol 被判断的单元格的列号
     * @return
     * @throws IOException
     */
    private static int getMergerCellRegionCol(HSSFSheet sheet, int cellRow,
                                              int cellCol) throws IOException {
        int retVal = 0;
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
            int firstRow = cra.getFirstRow(); // 合并单元格CELL起始行
            int firstCol = cra.getFirstColumn(); // 合并单元格CELL起始列
            int lastRow = cra.getLastRow(); // 合并单元格CELL结束行
            int lastCol = cra.getLastColumn(); // 合并单元格CELL结束列
            if (cellRow >= firstRow && cellRow <= lastRow) { // 判断该单元格是否是在合并单元格中
                if (cellCol >= firstCol && cellCol <= lastCol) {
                    retVal = lastCol - firstCol + 1; // 得到合并的列数
                    break;
                }
            }
        }
        return retVal;
    }

    /**
     * 判断单元格是否是合并的单格，如果是，获取其合并的行数。
     *
     * @param sheet   表单
     * @param cellRow 被判断的单元格的行号
     * @param cellCol 被判断的单元格的列号
     * @return
     * @throws IOException
     */
    private static int getMergerCellRegionRow(HSSFSheet sheet, int cellRow,
                                              int cellCol) throws IOException {
        int retVal = 0;
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
            int firstRow = cra.getFirstRow(); // 合并单元格CELL起始行
            int firstCol = cra.getFirstColumn(); // 合并单元格CELL起始列
            int lastRow = cra.getLastRow(); // 合并单元格CELL结束行
            int lastCol = cra.getLastColumn(); // 合并单元格CELL结束列
            if (cellRow >= firstRow && cellRow <= lastRow) { // 判断该单元格是否是在合并单元格中
                if (cellCol >= firstCol && cellCol <= lastCol) {
                    retVal = lastRow - firstRow + 1; // 得到合并的行数
                    break;
                }
            }
        }
        return 0;
    }

    /**
     * 单元格背景色转换
     *
     * @param hc
     * @return
     */
    private String convertToStardColor(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            int a = HSSFColor.AUTOMATIC.index;
            int b = hc.getIndex();
            if (a == b) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                String str;
                String str_tmp = Integer.toHexString(hc.getTriplet()[i]);
                if (str_tmp != null && str_tmp.length() < 2) {
                    str = "0" + str_tmp;
                } else {
                    str = str_tmp;
                }
                sb.append(str);
            }
        }
        return sb.toString();
    }

    /**
     * 单元格小平对齐
     *
     * @param alignment
     * @return
     */
    private String convertAlignToHtml(short alignment) {
        String align = "left";
        switch (alignment) {
            case HSSFCellStyle.ALIGN_LEFT:
                align = "left";
                break;
            case HSSFCellStyle.ALIGN_CENTER:
                align = "center";
                break;
            case HSSFCellStyle.ALIGN_RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * 单元格垂直对齐
     *
     * @param verticalAlignment
     * @return
     */
    private String convertVerticalAlignToHtml(short verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
            case HSSFCellStyle.VERTICAL_BOTTOM:
                valign = "bottom";
                break;
            case HSSFCellStyle.VERTICAL_CENTER:
                valign = "center";
                break;
            case HSSFCellStyle.VERTICAL_TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    public void makeFile() {
        String sdStateString = android.os.Environment.getExternalStorageState();// 获取外部存储状态
        if (sdStateString.equals(android.os.Environment.MEDIA_MOUNTED)) {// 确认sd卡存在,原理不知,媒体安装??
            try {
                File dirFile = new File(picPath);// 获取xiao文件夹地址
                if (!dirFile.exists()) {// 如果不存在
                    dirFile.mkdir();// 创建目录
                }
                File myFile = new File(picPath + File.separator + fileNameNoExtension + ".html");// 获取my.html的地址
                if (!myFile.exists()) {// 如果不存在
                    myFile.createNewFile();// 创建文件
                }
                this.htmlPath = myFile.getAbsolutePath();// 返回路径
            } catch (Exception e) {
            }
        }
    }

    /* 用来在sdcard上创建图片 */
    public void makePictureFile() {
        String sdString = android.os.Environment.getExternalStorageState();// 获取外部存储状态
        if (sdString.equals(android.os.Environment.MEDIA_MOUNTED)) {// 确认sd卡存在,原理不知
            try {
                File picDirFile = new File(picPath);
                if (!picDirFile.exists()) {
                    picDirFile.mkdir();
                }
                File pictureFile = new File(picPath + File.separator
                        + presentPicture + ".jpg");// 创建jpg文件,方法与html相同
                if (!pictureFile.exists()) {
                    pictureFile.createNewFile();
                }
                this.picturePath = pictureFile.getAbsolutePath();// 获取jpg文件绝对路径
            } catch (Exception e) {
                System.out.println("PictureFile Catch Exception");
            }
        }
    }

    public void writePicture() {
        Picture picture = (Picture) pictures.get(presentPicture);
        byte[] pictureBytes = picture.getContent();
        Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0,
                pictureBytes.length);
        makePictureFile();
        presentPicture++;
        File myPicture = new File(picturePath);
        try {
            FileOutputStream outputPicture = new FileOutputStream(myPicture);
            outputPicture.write(pictureBytes);
            outputPicture.close();
        } catch (Exception e) {
            System.out.println("outputPicture Exception");
        }
//        String imageString = "<img src=\"" + picturePath + "\"";
//        imageString = imageString + ">";
        String imageString = "<img src=\"./" + myPicture.getName() + "\"";
        imageString = imageString + ">";
        try {
            output.write(imageString.getBytes());
        } catch (Exception e) {
            System.out.println("output Exception");
        }
    }

    public int decideSize(int size) {

        if (size >= 1 && size <= 8) {
            return 1;
        }
        if (size >= 9 && size <= 11) {
            return 2;
        }
        if (size >= 12 && size <= 14) {
            return 3;
        }
        if (size >= 15 && size <= 19) {
            return 4;
        }
        if (size >= 20 && size <= 29) {
            return 5;
        }
        if (size >= 30 && size <= 39) {
            return 6;
        }
        if (size >= 40) {
            return 7;
        }
        return 3;
    }

    private String decideColor(int a) {
        int color = a;
        switch (color) {
            case 1:
                return "#000000";
            case 2:
                return "#0000FF";
            case 3:
            case 4:
                return "#00FF00";
            case 5:
            case 6:
                return "#FF0000";
            case 7:
                return "#FFFF00";
            case 8:
                return "#FFFFFF";
            case 9:
                return "#CCCCCC";
            case 10:
            case 11:
                return "#00FF00";
            case 12:
                return "#080808";
            case 13:
            case 14:
                return "#FFFF00";
            case 15:
                return "#CCCCCC";
            case 16:
                return "#080808";
            default:
                return "#000000";
        }
    }

    private void getRange() {
        FileInputStream in = null;
        POIFSFileSystem pfs = null;

        try {
            in = new FileInputStream(nameStr);
            pfs = new POIFSFileSystem(in);
            hwpf = new HWPFDocument(pfs);
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        range = hwpf.getRange();

        pictures = hwpf.getPicturesTable().getAllPictures();

        tableIterator = new TableIterator(range);

    }

    public void writeDOCXPicture(byte[] pictureBytes) {
        Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0,
                pictureBytes.length);
        makePictureFile();
        this.presentPicture++;
        File myPicture = new File(this.picturePath);
        try {
            FileOutputStream outputPicture = new FileOutputStream(myPicture);
            outputPicture.write(pictureBytes);
            outputPicture.close();
        } catch (Exception e) {
            System.out.println("outputPicture Exception");
        }
        String imageString = "<img src=\"" + this.picturePath + "\"";

        imageString = imageString + ">";
        try {
            this.output.write(imageString.getBytes());
        } catch (Exception e) {
            System.out.println("output Exception");
        }
    }

    public void writeParagraphContent(Paragraph paragraph) {
        Paragraph p = paragraph;
        int pnumCharacterRuns = p.numCharacterRuns();

        for (int j = 0; j < pnumCharacterRuns; j++) {

            CharacterRun run = p.getCharacterRun(j);

            if (run.getPicOffset() == 0 || run.getPicOffset() >= 1000) {
                if (presentPicture < pictures.size()) {
                    writePicture();
                }
            } else {
                try {
                    String text = run.text();
                    if (text.length() >= 2 && pnumCharacterRuns < 2) {
                        output.write(text.getBytes());
                    } else {
                        int size = run.getFontSize();
                        int color = run.getColor();
                        String fontSizeBegin = "<font size=\""
                                + decideSize(size) + "\">";
                        String fontColorBegin = "<font color=\""
                                + decideColor(color) + "\">";
                        String fontEnd = "</font>";
                        String boldBegin = "<b>";
                        String boldEnd = "</b>";
                        String islaBegin = "<i>";
                        String islaEnd = "</i>";

                        output.write(fontSizeBegin.getBytes());
                        output.write(fontColorBegin.getBytes());

                        if (run.isBold()) {
                            output.write(boldBegin.getBytes());
                        }
                        if (run.isItalic()) {
                            output.write(islaBegin.getBytes());
                        }

                        output.write(text.getBytes());

                        if (run.isBold()) {
                            output.write(boldEnd.getBytes());
                        }
                        if (run.isItalic()) {
                            output.write(islaEnd.getBytes());
                        }
                        output.write(fontEnd.getBytes());
                        output.write(fontEnd.getBytes());
                    }
                } catch (Exception e) {
                    System.out.println("Write File Exception");
                }
            }
        }
    }
}



