package com.ys.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
    public final static String path = "E:\\PPTpoi\\aqyd\\img\\";
    private static Map img;

    public static void main(String[] args) throws Exception {
        /*
         * 网页输出
         */
        FileOutputStream fs = new FileOutputStream(new File("E:\\PPTpoi\\aqyd\\output.html"));
        PrintStream printStream = new PrintStream(fs);
        printStream.println(
                "<!DOCTYPE html>\n<html lang=\"en\" xmlns=\"http://www.w3.org/1/xhtml\">\n<head>\n<meta charset=\"utf-8\" />" + "<title>js实现ppt</title>");

        // 加载PPT
        HSLFSlideShow ss = new HSLFSlideShow(new HSLFSlideShowImpl("E:\\PPTpoi\\aqyd\\《安全用电》教学课件.ppt"));
        img = new HashMap();

        // extract all pictures contained in the presentation
        int idx = 1;
        for (HSLFPictureData pict : ss.getPictureData()) {
            // picture data
            byte[] data = pict.getData();
            PictureData.PictureType type = pict.getType();
            String ext = type.extension;
            FileOutputStream out = new FileOutputStream(path + idx + ext);
            out.write(data);
            if(ext.equals(".wmf")){
                Wmf2Svg.convert(path + idx + ext);
                ext=".svg";
            }
            img.put(idx, idx + ext);
            out.close();
            idx++;
        }

        List<HSLFSlide> slides = ss.getSlides();
        printStream.println(JSswitch(slides.size()));
        for (HSLFSlide slide : slides) {
            List<List<HSLFTextParagraph>> textPss = slide.getTextParagraphs();
            for (List<HSLFTextParagraph> textPs : textPss) {
                // for (HSLFTextParagraph textP : textPss) {
                List<HSLFTextRun> trs = textPs.get(0).getTextRuns();
                for (HSLFTextRun tr : trs) {
//                    System.out.println("text:    " + tr.getRawText() + "  " + tr.getFontFamily() + "  " + tr.getFontIndex() + "  " + tr.getFontSize() + "  "
//                            + tr.getFontColor().getSolidColor().getColor().toString());
                }
                // }
            }
        }
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
                System.out.println("框类型"+shape.getClass().toGenericString()+"         "+shape.getShapeName()+"         ");

                if (shape instanceof HSLFTextShape) {
                    HSLFTextShape tsh = (HSLFTextShape) shape;
                    // 获取关于文字更加详细的信息
                    System.out.println("文字和位置：    " + tsh.getText() + "    " + tsh.getFillColor() + "   " + tsh.getAnchor().toString());

                } else if (shape instanceof HSLFPictureShape) {
                    HSLFPictureShape tsh = (HSLFPictureShape) shape;
                    System.out.println("读取图片 ： " + tsh.getAnchor() + "     picNum: " + tsh.getPictureIndex());
                    printStream.println(insertImg(j, tsh.getPictureIndex(), tsh.getAnchor().toString()));
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
                + "px;left: " + x + "px;width: " + w + "px;height: " + h + "px;\")" + ";image" + i + ".src = \"img/" + img.get(imgIndex)+"\";"
                + "div_all.appendChild(image" + i + ");";
        // System.out.println(result);
        return result;
    }

    // js标签内，向div插入文字，参数文字几，文字，绝对位置，文字字体，文字颜色，文字宽高，文本框背景色
    public static String inserP(int j, String str, String pos) {
        String all = pos.substring(pos.indexOf('[') + 1, pos.indexOf(']'));
        String[] alls, xs, ys, ws, hs;
        alls = all.split(",");
        xs = alls[0].split("=");
        ys = alls[1].split("=");
        ws = alls[2].split("=");
        hs = alls[3].split("=");
        int x, y, h, w;
        x = Integer.valueOf(xs[1]);
        y = Integer.valueOf(ys[1]);
        h = Integer.valueOf(hs[1]);
        w = Integer.valueOf(ws[1]);
        String result = "\n var p" + j + " = document.createElement(\"p\");" + "\n p" + j + ".setAttribute(\"style\", \"position: absolute;top: " + y
                + "px;left: " + x + "px;width: " + w + "px;height: " + h + "px;\")" + ";p" + j + ".innerHTML =\" " + str + "\";\n" + "div_all.appendChild(p" + j
                + ");";
        //System.out.println(result);
        return result;
    }

}
