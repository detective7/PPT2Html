package com.ys.util;

import java.io.PrintStream;

import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.PaintStyle.SolidPaint;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PPTXhtmlUtil {

    private PrintStream printStream;

    public PPTXhtmlUtil(PrintStream printStream) {
        this.printStream = printStream;
    }

    public void start() {
        printStream.println(
                "<!DOCTYPE html>\n<html lang=\"en\" xmlns=\"http://www.w3.org/1/xhtml\">\n<head>\n<meta charset=\"utf-8\" />" + "<title>js实现ppt</title>");
    }

    // js标签开始
    public void JSstart(int num) {
        StringBuilder result =  new StringBuilder("\n<script type=\"text/javascript\" language=\"JavaScript\">\n var cou = 0;function selAll() {switch(cou) {");
        for (int i = 0; i < num; i++) {
            result.append( "case " + i + ":fun" + i + "();cou++;break;");
        }
        result.append("}}");
        printStream.println(result);
    }

    public void end(double d, double e) {
        printStream.println(
                "\n</script>\n\n<style type=\"text/css\">\n body {\nmargin: 0;\npadding: 0;\n}\n</style>\n</head>\n<body>\n<div id=\"all\" onclick=\"selAll()\" style=\"width:"
                        + d + "px;height:" + e + "px;border: 1px solid;\">" + "\n<p id=\" text \"></p>" + "\n</div>\n</body>\n</html>");
    }

    // js标签内，向div插入图片，参数图片几，图片名，图片绝对位置
    public void insertImg(String imgName, String pos) {
        //System.out.println(pos);
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

        StringBuilder result = new StringBuilder("\n var image" + " = document.createElement(\"img\");" + "\n image" + ".setAttribute(\"style\", \"position: absolute;top: " + y
                + "px;left: " + x + "px;width: " + w + "px;height: " + h + "px;\")" + ";image" + ".src = \"imgPPTX/" + imgName + "\";"
                + "div_all.appendChild(image" + ");");
        // System.out.println(result);
        printStream.println(result);
    }

    // js标签内，插入文字标签p，绝对位置，宽高，文本框背景色
    public void insertP(int j, XSLFTextShape txShape) {
        String pos=txShape.getAnchor().toString();
        //System.out.println(pos);
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
        StringBuilder result =  new StringBuilder("");
        if (txShape.getFillColor() != null) {
            String bgColor=TransformUtil.toHex(txShape.getFillColor().toString());
            double bgAlpha=(double)txShape.getFillColor().getAlpha()/255d;
            result.append("\n var p" + j + " = document.createElement(\"p\");" + "\n p" + j + ".setAttribute(\"style\", \"background:"+bgColor+";margin:0;opacity:"+bgAlpha+";position: absolute;top: " + y + "px;left: "
                    + x + "px;width: " + w + "px;height: " + h + "px;\");\n" + "div_all.appendChild(p" + j + ");");
        } else {
            result.append("\n var p" + j + " = document.createElement(\"p\");" + "\n p" + j + ".setAttribute(\"style\", \"position: absolute;top: " + y + "px;left: "
                    + x + "px;width: " + w + "px;height: " + h + "px;\")" + ";\n" + "div_all.appendChild(p" + j + ");");
        }
        // System.out.println(result);
        printStream.println(result);
    }

    // 向p里添加span标签，输入文字
    public void insertSpan(int j, XSLFTextRun text, String str) {
        StringBuilder result = new StringBuilder("\n var span = document.createElement(\"span\");");
        if (text != null) {
            if (text.getFontColor() instanceof PaintStyle.SolidPaint) {
                SolidPaint color = (PaintStyle.SolidPaint) text.getFontColor();
                result.append("\n span.setAttribute(\"style\", \"font-family:" + text.getFontFamily() + ";font-size: " + text.getFontSize() + "px;color:"
                        + TransformUtil.toHex(color.getSolidColor().getColor().toString()));
            } else {
                result.append("\n span.setAttribute(\"style\", \"font-family:" + text.getFontFamily() + ";font-size: " + text.getFontSize() + "px");
            }
            if(text.isBold()){
                result.append("; font-weight: bold");
            }
            if(text.isItalic()){
                result.append(";font-style: italic");
            }
            if(text.isUnderlined()){
                result.append("; text-decoration: underline");
            }
            result.append(";margin: 0;"+ "\");\nspan.innerHTML = \"" + str + "\";\n" + "p" + j + ".appendChild(span);");
        } else {
            result.append("\nspan.innerHTML = \"" + str + "\";\n" + "p" + j + ".appendChild(span);");
        }
        // System.out.println(str);
        printStream.println(result);
    }

    public void insertDiv(XSLFGroupShape group) {
        String pos=group.getAnchor().toString();
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
        StringBuilder result = new StringBuilder("\n var div"+group.getShapeId()+" = document.createElement(\"div\");\n");
        result.append("div"+group.getShapeId()+".setAttribute(\"style\", \"width:"+w+"px;height:"+h+"px;border: 2px solid;left: "+x+"px;top: "+y+"px;position:absolute;\");");
        result.append("\n div_all.appendChild(div"+group.getShapeId()+ ");");
        printStream.println(result);
    }
    
}
