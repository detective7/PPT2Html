package com.ys.poi;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackground;

import com.ys.util.PPTXhtmlUtil;
import com.ys.util.Wmf2Svg;

public class PPTXpoi315 {

    // 图片存放路径
    private static Map<String, String> img;
    private static FileOutputStream fs;
    private static PrintStream printStream;
    private static String path;
    private static PPTXhtmlUtil htmlins;
    private static Dimension pageSize;
    //p计数量
//    private static int shapeNum;

    public static void main(String[] args) throws Exception {
        
        path = "E:\\PPTpoi\\y\\";
        FileInputStream is = new FileInputStream(path + "yx.pptx");
        XMLSlideShow ppts = new XMLSlideShow(is);
        is.close();
        img = new HashMap<String, String>();
        pageSize = ppts.getPageSize();
//        shapeNum=0;

        fs = new FileOutputStream(new File(path + "outpptx.html"));
        printStream = new PrintStream(fs);

        getPic(ppts);
        //System.out.println(img.toString());

        htmlins = new PPTXhtmlUtil(printStream);
        htmlins.start();
        htmlins.JSstart(ppts.getSlides().size());

        List<XSLFSlide> slides = ppts.getSlides();

        for (int i = 0; i < slides.size(); i++) {
            /*
             * 每个界面的js函数
             */
            printStream.println("\n function fun" + i
                    + "() {\n var div_all = document.getElementById(\"all\");\n if(div_all) {while(div_all.hasChildNodes()) {div_all.removeChild(div_all.firstChild);}");

            // 获取背景，母版，主题和布局，主要是图片，布局的文字还没添加
            getBg(slides.get(i));

            // 获取文字和图片
            for (XSLFShape shape : slides.get(i)) {
                dealGroupShape(shape,null);
            }
            printStream.println("}}");
        }
        // printStream.println(ppts.getSlides().get(1).getSlideLayout().getBackground().getXmlObject());//.getXmlObject());//.getTheme().getXmlObject().xmlText());
        
        htmlins.end(pageSize.getWidth(), pageSize.getHeight());

        printStream.close();
    }

    // 输出图片
    public static void getPic(XMLSlideShow ppts) throws Exception {
        // for (PackagePart p : ppts.getAllEmbedds()) {
        // String type = p.getContentType();
        // String name = p.getPartName().getName();
        // //out.println("Embedded file (" + type + "): " + name);
        // InputStream pIs = p.getInputStream();
        // pIs.close();
        // }

        for (XSLFPictureData data : ppts.getPictureData()) {
            String index = data.getIndex()+1 + "";
            String ext = data.getType().extension;

            FileOutputStream fileout = new FileOutputStream(path + "imgPPTX\\" + index + ext);
            InputStream pIs = data.getInputStream();
            if (ext.equals(".wmf")) {
                Wmf2Svg.convert(path + index + ext);
                ext = ".svg";
            }
            // System.out.println(pict.getHeader().toString());
            img.put(data.getFileName(), index + ext);

            fileout.write(data.getData());
            fileout.close();
            pIs.close();
        }

    }

    public static void getBg(XSLFSlide ppt) throws Exception {
        System.out.println(ppt.getSlideNumber() + ":  " + ppt.getSlideLayout().getTheme().getName() + "   " + ppt.getSlideLayout().getName());
        // .println(ppt.getSlideLayout().getBackground().getXmlObject());
        // System.out.println(ppt.getSlideNumber() + ": " +
        // ppt.getBackground().getXmlObject());
        XSLFBackground bg = ppt.getBackground();
        CTBackground xmlBg = (CTBackground) bg.getXmlObject();
        if (xmlBg.getBgPr().getBlipFill() != null) {
            String relId = xmlBg.getBgPr().getBlipFill().getBlip().getEmbed();

            XSLFPictureData pic = (XSLFPictureData) ppt.getRelationById(relId);
            htmlins.insertImg(img.get(pic.getFileName()), "[x=0,y=0,w="+pageSize.getWidth()+",h="+pageSize.getHeight()+"]");
            // System.out.println("backg: " + pic.getFileName());
        } else {
            // System.out.println("backg: 没背景");
        }
        for (XSLFShape shape : ppt.getSlideLayout().getShapes()) {
                dealBGShape(shape);
        }
    }

    public static void dealGroupShape(XSLFShape shape,String shapeId) {
        if (shape instanceof XSLFGroupShape) {
            XSLFGroupShape group = (XSLFGroupShape) shape;
            htmlins.insertDiv(group);
            for (XSLFShape subShape : group.getShapes()) {
                dealGroupShape(subShape,group.getShapeId()+"");
            }
        } else if (shape instanceof XSLFTextShape) {
            XSLFTextShape txShape = (XSLFTextShape) shape;

            String txStr = txShape.getText().replaceAll("\n", "<br>");
            if (txShape.getFillColor() != null) {
//                shapeNum++;
                System.out.println(txShape.getVerticalAlignment().toString());
                htmlins.insertP(shape.getShapeId(), txShape);//txShape.getAnchor().toString(), TransformUtil.toHex(txShape.getFillColor().toString()));
            } else {
//                shapeNum++;
                htmlins.insertP(shape.getShapeId(), txShape);//.getAnchor().toString(), null);
            }
//            System.out.println("txStr: =>" + txStr+"   "+txShape.getAnchor().toString()+"   "+txShape.getFillColor());
            for (XSLFTextParagraph p : txShape) {
                // out.println("Paragraph level: " +
                // p.getIndentLevel());
                for (XSLFTextRun r : p) {
                    String reStr = r.getRawText().replaceAll("\n", "<br>");
                    for (; txStr.startsWith("<br>");) {
                        reStr = reStr + "<br>";
                        txStr = txStr.replaceFirst("<br>", "");
                    }
                    // System.out.println("deStr: ->" + reStr);
                    if (reStr.equals("[?]") || reStr.startsWith("?")) {
                        txStr = txStr.replaceFirst("\\" + reStr, "");
                    } else {
                        txStr = txStr.replaceFirst(reStr, "");
                    }
                    for (; txStr.startsWith(" ");) {
                        txStr = txStr.replaceFirst(" ", "");
                    }
                    // 回车换行得额外弄，按顺序把字符替换掉，剩下的，是上面textparam查不出的字符，再贴上去
                    for (; txStr.startsWith("<br>");) {
                        reStr = reStr + "<br>";
                        txStr = txStr.replaceFirst("<br>", "");
                    }
                    htmlins.insertSpan(txShape.getShapeId(), r, reStr);
//                    System.out.println(reStr+"   "+p.getTextAlign().toString());
//                    System.out.println(" font.family: " + r.getFontFamily());
                }
            }
        } else if (shape instanceof XSLFPictureShape) {
            XSLFPictureShape tsh = (XSLFPictureShape) shape;
            htmlins.insertImg(img.get(tsh.getPictureData().getFileName()), tsh.getAnchor().toString());
            // System.out.println("layout: " +
            // tsh.getPictureData().getFileName() + " " +
            // tsh.getAnchor().toString());
        } else {
//            System.out.println("识别不出: "+shape.getClass());
        }
    }
    
    public static void dealBGShape(XSLFShape shape) {
        if (shape instanceof XSLFGroupShape) {
            XSLFGroupShape group = (XSLFGroupShape) shape;
            for (XSLFShape subShape : group.getShapes()) {
                dealGroupShape(subShape,subShape.getShapeId()+"");
            }
        } else if (shape instanceof XSLFPictureShape) {
            XSLFPictureShape tsh = (XSLFPictureShape) shape;
            htmlins.insertImg(img.get(tsh.getPictureData().getFileName()), tsh.getAnchor().toString());
            // System.out.println("layout: " +
            // tsh.getPictureData().getFileName() + " " +
            // tsh.getAnchor().toString());
        } else {
            // System.out.println("layout: "+shape.getClass());
        }
    }
}
