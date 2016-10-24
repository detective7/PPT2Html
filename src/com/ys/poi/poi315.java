package com.ys.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.text.html.HTMLEditorKit.InsertHTMLTextAction;

import org.apache.poi.hslf.model.MovieShape;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.hslf.usermodel.HSLFTextShape;
import org.apache.poi.sl.usermodel.PictureData;

import com.ys.util.Wmf2Svg;

public class poi315 {

    // 图片默认存放路径
    public final static String path = "E:\\PPTpoi\\安全用电\\img\\";
    private static Map img;
    private static List<PPTText> texts;

    public static void main(String[] args) throws Exception {
        /*
         * 网页输出
         */
        FileOutputStream fs = new FileOutputStream(new File("E:\\PPTpoi\\安全用电\\output.html"));
        PrintStream printStream = new PrintStream(fs);
        printStream.println(
                "<!DOCTYPE html>\n<html lang=\"en\" xmlns=\"http://www.w3.org/1/xhtml\">\n<head>\n<meta charset=\"utf-8\" />" + "<title>js实现ppt</title>");

        // 加载PPT
        HSLFSlideShow ss = new HSLFSlideShow(new HSLFSlideShowImpl("E:\\PPTpoi\\安全用电\\《安全用电》教学课件.ppt"));
        img = new HashMap();

        // 存字相关信息
        texts = new ArrayList<PPTText>();

        // 取所有图片，并把矢量图wmf转为网页可显示的svg格式。 extract all pictures contained in the
        // presentation
        int idx = 1;
        for (HSLFPictureData pict : ss.getPictureData()) {
            // picture data
            byte[] data = pict.getData();
            PictureData.PictureType type = pict.getType();
            String ext = type.extension;
            FileOutputStream out = new FileOutputStream(path + idx + ext);
            out.write(data);
            if (ext.equals(".wmf")) {
                Wmf2Svg.convert(path + idx + ext);
                ext = ".svg";
            }
            img.put(idx, idx + ext);
            out.close();
            idx++;
        }

        List<HSLFSlide> slides = ss.getSlides();
        printStream.println(JSswitch(slides.size()));
        for (HSLFSlide slide : slides) {
            System.out.println(slide.getSlideNumber() + "     PPT");
            List<List<HSLFTextParagraph>> textPss = slide.getTextParagraphs();
            for (List<HSLFTextParagraph> textPs : textPss) {
                for (HSLFTextParagraph textP : textPs) {
                    List<HSLFTextRun> trs = textP.getTextRuns();
                    for (HSLFTextRun tr : trs) {
                        PPTText text = new PPTText();
                        String t = tr.getRawText().replaceAll("", "");
                        if (t != null && !t.equals("") && !t.trim().equals("")) {
                            text.setSlideNum(slide.getSlideNumber());
                            text.setText(t);
                            text.setColor(toHex(tr.getFontColor().getSolidColor().getColor().toString()));
                            text.setSize(tr.getFontSize());
                            text.setFontFamily(tr.getFontFamily());
                            texts.add(text);
//                            System.out.println("text: " + text.getText() + " " + tr.getFontFamily() + " " + tr.getFontSize() + " "
//                                    + toHex(tr.getFontColor().getSolidColor().getColor().toString()));
                        }
                    }
                }
            }
        }
        
        System.out.println(texts);
        /**
         * 插入总的switch语句
         */
        for (int i = 0; i < slides.size(); i++) {
            /*
             * 每个界面的js函数
             */
            printStream.println("\n function fun" + i
                    + "() {\n var div_all = document.getElementById(\"all\");\n if(div_all) {while(div_all.hasChildNodes()) {div_all.removeChild(div_all.firstChild);}");

            int j = 0;
            for (HSLFShape shape : slides.get(i).getShapes()) {
                // System.out.println("框类型" + shape.getClass().toGenericString()
                // + " " + shape.getShapeName() + " ");
                // System.out.println(shape.getClass().toString() + " " +
                // MovieShape.class.toString());
                if (shape instanceof HSLFTextShape) {
                    HSLFTextShape tsh = (HSLFTextShape) shape;
                    // 获取关于文字更加详细的信息
                    if (tsh.getFillColor() != null) {
                        printStream.println(insertP(j, tsh.getAnchor().toString(), toHex(tsh.getFillColor().toString())));
                    } else {
                        printStream.println(insertP(j, tsh.getAnchor().toString(), null));
                    }
                    for (int textNum = 0; textNum < texts.size(); textNum++) {
                        PPTText t = texts.get(textNum);
                        if (slides.get(i).getSlideNumber() == t.getSlideNum() && tsh.getText().contains(t.getText())) {
                            System.out.println(texts.get(textNum).getText());
                            printStream.println(insertSpan(j, t));
                        }
                    }
                } else if (shape instanceof HSLFPictureShape) {
                    HSLFPictureShape tsh = (HSLFPictureShape) shape;
                    // System.out.println("读取图片 ： " + tsh.getAnchor() + "
                    // picNum: " + tsh.getPictureIndex());
                    printStream.println(insertImg(j, tsh.getPictureIndex(), tsh.getAnchor().toString()));
                    if (shape.getClass().toString().equals(MovieShape.class.toString())) {
                        MovieShape ms = (MovieShape) shape;
                        // System.out.println("视频音频： " + ms.getPath());
                    }
                }
                /* System.out.println("Shape     " + shape.getClass()); */
                j++;
            }
            System.out.println("第" + slides.get(i).getSlideNumber() + "张PPT解析结束  \n");
            printStream.println("}}");
        }
        printStream.println(endDiv(ss.getPageSize().width, ss.getPageSize().height));
        printStream.close();
    }

    // js标签开始
    public static String JSswitch(int num) {
        String result = "\n\n\n\n\n<script type=\"text/javascript\" language=\"JavaScript\">\n var cou = 0;function selAll() {switch(cou) {";
        for (int i = 0; i < num; i++) {
            result = result + "case " + i + ":fun" + i + "();cou++;break;";
        }
        result += "}}";
        return result;
    }

    // 添加div标签，主要设置div的宽高
    public static String endDiv(int width, int height) {
        return "\n\n\n\n\n</script>\n<style type=\"text/css\">\n body {\nmargin: 0;\npadding: 0;\n}\n</style>\n</head>\n<body>\n<div id=\"all\" onclick=\"selAll()\" style=\"width:"
                + width + "px;height:" + height + "px;border: 1px solid;\">" + "\n<p id=\" text \"></p>" + "\n</div>\n</body>\n</html>";
    }

    // js标签内，向div插入图片，参数图片几，图片名，图片绝对位置
    public static String insertImg(int i, int imgIndex, String pos) {
        String all = pos.substring(pos.indexOf('[') + 1, pos.indexOf(']'));
        String[] alls, xs, ys, ws, hs;
        alls = all.split(",");
        xs = alls[0].split("=");
        ys = alls[1].split("=");
        ws = alls[2].split("=");
        hs = alls[3].split("=");
        float x, y, h, w;
        x = Float.valueOf(xs[1]);
        y = Float.valueOf(ys[1]);
        h = Float.valueOf(hs[1]);
        w = Float.valueOf(ws[1]);

        String result = "\n var image" + i + " = document.createElement(\"img\");" + "\n image" + i + ".setAttribute(\"style\", \"position: absolute;top: " + y
                + "px;left: " + x + "px;width: " + w + "px;height: " + h + "px;\")" + ";image" + i + ".src = \"img/" + img.get(imgIndex) + "\";"
                + "div_all.appendChild(image" + i + ");";
        // System.out.println(result);
        return result;
    }

    // js标签内，插入文字div，绝对位置，宽高，文本框背景色
    public static String insertP(int j, String pos, String bgColor) {
        String all = pos.substring(pos.indexOf('[') + 1, pos.indexOf(']'));
        String[] alls, xs, ys, ws, hs;
        alls = all.split(",");
        xs = alls[0].split("=");
        ys = alls[1].split("=");
        ws = alls[2].split("=");
        hs = alls[3].split("=");
        float x, y, h, w;
        x = Float.valueOf(xs[1]);
        y = Float.valueOf(ys[1]);
        h = Float.valueOf(hs[1]);
        w = Float.valueOf(ws[1]);
        String result = null;
        if (bgColor != null) {
            result = "\n var p" + j + " = document.createElement(\"p\");" + "\n p" + j + ".setAttribute(\"style\", \"position: absolute;top: " + y + "px;left: "
                    + x + "px;width: " + w + "px;height: " + h + "px;background:" + bgColor + ";\");\n" + "div_all.appendChild(p" + j + ");";
        } else {
            result = "\n var p" + j + " = document.createElement(\"p\");" + "\n p" + j + ".setAttribute(\"style\", \"position: absolute;top: " + y + "px;left: "
                    + x + "px;width: " + w + "px;height: " + h + "px;\")" + ";\n" + "div_all.appendChild(p" + j + ");";
        }
        // System.out.println(result);
        return result;
    }

    // 向div里添加p标签，输入文字
    public static String insertSpan(int j, PPTText text) {
        String result = "\n var span = document.createElement(\"span\");" + "\n span.setAttribute(\"style\", \"font-family:" + text.getFontFamily()
                + ";font-size: " + text.getSize() + "px;color:" + text.getColor() + ";margin: 0;" + "\");\nspan.innerHTML = \"" + text.getText() + "\";\n" + "p"
                + j + ".appendChild(span);";
        // System.out.println(result);
        return result;
    }

    // RGB转16进制色
    public static String toHex(String rgb) {
        String all = rgb.substring(rgb.indexOf('[') + 1, rgb.indexOf(']'));
        String[] alls, rs, gs, bs;
        alls = all.split(",");
        rs = alls[0].split("=");
        gs = alls[1].split("=");
        bs = alls[2].split("=");
        int r, g, b;
        r = Integer.valueOf(rs[1]);
        g = Integer.valueOf(gs[1]);
        b = Integer.valueOf(bs[1]);

        return "#" + toBrowserHexValue(r) + toBrowserHexValue(g) + toBrowserHexValue(b);
    }

    private static String toBrowserHexValue(int number) {
        StringBuilder builder = new StringBuilder(Integer.toHexString(number & 0xff));
        while (builder.length() < 2) {
            builder.append("0");
        }
        return builder.toString().toUpperCase();
    }

}
