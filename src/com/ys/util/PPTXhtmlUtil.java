package com.ys.util;

import java.io.PrintStream;
import java.util.Map;

public class PPTXhtmlUtil {
    
    private static PrintStream printStream;
    private static Map img;
    
    public PPTXhtmlUtil(PrintStream printStream){
        this.printStream = printStream;
    }
    
    public static void start(){
        printStream.println(
                "<!DOCTYPE html>\n<html lang=\"en\" xmlns=\"http://www.w3.org/1/xhtml\">\n<head>\n<meta charset=\"utf-8\" />" + "<title>js实现ppt</title>");
    }
    
    // js标签开始
    public static void JSstart(int num) {
        String result = "\n<script type=\"text/javascript\" language=\"JavaScript\">\n var cou = 0;function selAll() {switch(cou) {";
        for (int i = 0; i < num; i++) {
            result = result + "case " + i + ":fun" + i + "();cou++;break;";
        }
        result += "}}";
        printStream.println(result);
    }
    
    public static void end(double d, double e) {
        printStream.println("\n</script>\n\n<style type=\"text/css\">\n body {\nmargin: 0;\npadding: 0;\n}\n</style>\n</head>\n<body>\n<div id=\"all\" onclick=\"selAll()\" style=\"width:"
                + d + "px;height:" + e + "px;border: 1px solid;\">" + "\n<p id=\" text \"></p>" + "\n</div>\n</body>\n</html>");
    }
    
 // js标签内，向div插入图片，参数图片几，图片名，图片绝对位置
    public static void insertImg(String imgName, String pos) {
        System.out.println(pos);
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

        String result = "\n var image" + " = document.createElement(\"img\");" + "\n image" + ".setAttribute(\"style\", \"position: absolute;top: " + y
                + "px;left: " + x 
                + "px;width: " + w 
                + "px;height: " + h 
                + "px;\")" + ";image" 
                + ".src = \"imgPPTX/" + imgName + "\";"
                + "div_all.appendChild(image" + ");";
        // System.out.println(result);
        printStream.println(result);
    }
    
    

}
