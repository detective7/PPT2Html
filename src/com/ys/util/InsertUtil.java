package com.ys.util;

import java.io.PrintStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.hslf.model.MovieShape;
import org.apache.poi.hslf.usermodel.HSLFGroupShape;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFTextShape;

import com.ys.poi.PPTText;

public class InsertUtil {

    private static PrintStream printStream;
    private static Map img;
    private static List<PPTText> texts;

    public InsertUtil(PrintStream printStream, List<PPTText> texts, Map img) {
        this.printStream = printStream;
        this.img = img;
        this.texts = texts;
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
    public static String insertSpan(int j, PPTText text, String str) {
        String result = "\n var span = document.createElement(\"span\");";
        if (text != null) {
            result += "\n span.setAttribute(\"style\", \"font-family:" + text.getFontFamily() + ";font-size: " + text.getSize() + "px;color:" + text.getColor()
                    + ";margin: 0;" + "\");\nspan.innerHTML = \"" + str + "\";\n" + "p" + j + ".appendChild(span);";
        } else {
            result += "\nspan.innerHTML = \"" + str + "\";\n" + "p" + j + ".appendChild(span);";
        }
        // System.out.println(str);
        return result;
    }


    // 用于递归处理GroupShape
    public static void dealWithGroup(int j, int slideNumber, HSLFShape shape) {
        // System.out.println(shape.getShapeId());
        if (shape instanceof HSLFTextShape) {
            HSLFTextShape tsh = (HSLFTextShape) shape;
            // 获取关于文字更加详细的信息
            if (tsh.getFillColor() != null) {
                printStream.println(insertP(j, tsh.getAnchor().toString(), TransformUtil.toHex(tsh.getFillColor().toString())));
            } else {
                printStream.println(insertP(j, tsh.getAnchor().toString(), null));
            }
            String shapeText = tsh.getText().toString();
            for (int textNum = 0; textNum < texts.size(); textNum++) {
                PPTText t = texts.get(textNum);

                if (slideNumber == t.getSlideNum() && shapeText.contains(t.getText().trim())) {
                    String insertT = t.getText().toString();
                    // 防止出现Dangling meta character '?' near index 0 ? Can you
                    // find them.这类错误
                    // System.out.println(slideNumber+"
                    // =>"+shapeText.replaceAll("\n", " 换行"));
                    // System.out.println("填充: ->"+insertT);
                    if (insertT.equals("[?]") || insertT.startsWith("?")) {
                        shapeText = shapeText.replaceFirst("\\" + insertT, "");
                    } else {
                        shapeText = shapeText.replaceFirst(insertT, "");
                    }
                    for (; shapeText.startsWith(" ");) {
                        shapeText = shapeText.replaceFirst(" ", "");
                    }
                    // 回车换行得额外弄，按顺序把字符替换掉,剩下的，是上面textparam查不出的字符，再贴上去
                    if (shapeText.startsWith("\n")) {
                        insertT = insertT + "<br>";
                        shapeText = shapeText.replaceFirst("\n", "");
                    }
                    // System.out.println("shapeText: "+shapeText);
                    printStream.println(insertSpan(j, t, insertT));
                }
            }
            if (!shapeText.trim().isEmpty()) {
                // System.out.println("还没完，"+shapeText.trim());
                printStream.println(insertSpan(j, null, shapeText.replace("\n", "<br>")));
            }
        } else if (shape instanceof HSLFPictureShape) {
            HSLFPictureShape tsh = (HSLFPictureShape) shape;
            // System.out.println("读取图片 ： " + tsh.getAnchor() + "
            // picNum: " + tsh.getPictureIndex());
            printStream.println(insertImg(j, tsh.getPictureIndex(), tsh.getAnchor().toString()));
            if (shape instanceof MovieShape) {
                MovieShape ms = (MovieShape) shape;
                // System.out.println("视频音频： " + ms.getPath());
            }
        } else if (shape instanceof HSLFGroupShape) {
            HSLFGroupShape group = (HSLFGroupShape) shape;
            for (HSLFShape gs : group.getShapes()) {
                dealWithGroup(j, slideNumber, gs);
            }
        }
    }
}
