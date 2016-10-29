package com.ys.poi;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hslf.usermodel.HSLFShape;
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

import com.ys.util.Wmf2Svg;

public class PPTXpoi315 {

    // 图片存放路径
    private static Map img;
    private static FileOutputStream fs;
    private static PrintStream printStream;
    private static String path;

    public static void main(String[] args) throws Exception {
        path = "E:\\PPTpoi\\y\\";
        FileInputStream is = new FileInputStream(path + "yx.pptx");
        XMLSlideShow ppts = new XMLSlideShow(is);
        is.close();
        img = new HashMap();

        fs = new FileOutputStream(new File(path + "out.html"));
        printStream = new PrintStream(fs);

        // 获取背景，母版，主题和布局，主要是图片，布局的文字还没添加
        getBg(ppts);

        // 获取文字和图片
        // getTextPic(ppts);
        // printStream.println(ppts.getSlides().get(1).getSlideLayout().getBackground().getXmlObject());//.getXmlObject());//.getTheme().getXmlObject().xmlText());
        printStream.close();
    }

    // 获取文字图片
    public static void getTextPic(XMLSlideShow ppts) throws Exception {
        // for (PackagePart p : ppts.getAllEmbedds()) {
        // String type = p.getContentType();
        // String name = p.getPartName().getName();
        // //out.println("Embedded file (" + type + "): " + name);
        // InputStream pIs = p.getInputStream();
        // pIs.close();
        // }

        for (XSLFPictureData data : ppts.getPictureData()) {
            String name = data.getFileName();
            String ext = data.getType().extension;

            FileOutputStream fileout = new FileOutputStream(path + "imgPPTX\\" + name);
            InputStream pIs = data.getInputStream();
            System.out.println(ext);
            if (ext.equals(".wmf")) {
                Wmf2Svg.convert(path + data.getFileName());
                ext = ".svg";
            }
            // System.out.println(pict.getHeader().toString());
            img.put(data.getIndex(), data.getFileName());
            fileout.write(data.getData());
            fileout.close();
            pIs.close();
        }

        Dimension pageSize = ppts.getPageSize();
        // out.println("Pagesize: " + pageSize);

        for (XSLFSlide slide : ppts.getSlides()) {
            for (XSLFShape shape : slide) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape txShape = (XSLFTextShape) shape;

                    String txStr = txShape.getText().replaceAll("\n", "<br>");
                    System.out.println("txStr:   =>" + txStr);
                    for (XSLFTextParagraph p : txShape) {
                        // out.println("Paragraph level: " +
                        // p.getIndentLevel());
                        for (XSLFTextRun r : p) {
                            String reStr = r.getRawText().replaceAll("\n", "<br>");
                            for (; txStr.startsWith("<br>");) {
                                reStr = reStr + "<br>";
                                txStr = txStr.replaceFirst("<br>", "");
                            }
                            System.out.println("deStr:   ->" + reStr);
                            if (reStr.equals("[?]") || reStr.startsWith("?")) {
                                txStr = txStr.replaceFirst("\\" + reStr, "");
                            } else {
                                txStr = txStr.replaceFirst(reStr, "");
                            }
                            for (; txStr.startsWith(" ");) {
                                txStr = txStr.replaceFirst(" ", "");
                            }
                            // 回车换行得额外弄，按顺序把字符替换掉,剩下的，是上面textparam查不出的字符，再贴上去
                            for (; txStr.startsWith("<br>");) {
                                reStr = reStr + "<br>";
                                txStr = txStr.replaceFirst("<br>", "");
                            }
                            System.out.println("txStr:   =>" + txStr);
                            // System.out.println(" bold: " + r.isBold());
                            // System.out.println(" italic: " + r.isItalic());
                            // System.out.println(" underline: " +
                            // r.isUnderlined());
                            // System.out.println(" font.family: " +
                            // r.getFontFamily());
                            // System.out.println(" font.size: " +
                            // r.getFontSize());
                            // System.out.println(" font.color: " +
                            // r.getFontColor());
                        }
                    }
                } else if (shape instanceof XSLFPictureShape) {
                    XSLFPictureShape pShape = (XSLFPictureShape) shape;
                    XSLFPictureData pData = pShape.getPictureData();
                    // out.println(pData.getFileName());
                } else {
                    // out.println("Process me: " + shape.getClass());
                }
            }

        }
    }

    public static void getBg(XMLSlideShow ppts) throws Exception {
        for (XSLFSlide ppt : ppts.getSlides()) {
            System.out.println(ppt.getSlideNumber() + ":  " + ppt.getSlideLayout().getTheme().getName() + "   " + ppt.getSlideLayout().getName());
            // .println(ppt.getSlideLayout().getBackground().getXmlObject());
            // System.out.println(ppt.getSlideNumber() + ": " +
            // ppt.getBackground().getXmlObject());
            XSLFBackground bg = ppt.getBackground();
            CTBackground xmlBg = (CTBackground) bg.getXmlObject();
            if (xmlBg.getBgPr().getBlipFill() != null) {
                String relId = xmlBg.getBgPr().getBlipFill().getBlip().getEmbed();

                XSLFPictureData pic = (XSLFPictureData) ppt.getRelationById(relId);
                System.out.println("backg:   " + pic.getFileName());
            } else {
                System.out.println("backg:   没背景");
            }
            for (XSLFShape shape : ppt.getSlideLayout().getShapes()) {
                dealGroupShape(shape);
            }
        }
    }

    public static void dealGroupShape(XSLFShape shape) {
        if (shape instanceof XSLFPictureShape) {
            XSLFPictureShape tsh = (XSLFPictureShape) shape;
            System.out.println("layout:   " + tsh.getPictureData().getFileName() + "   " + tsh.getAnchor().toString());
        } else if (shape instanceof XSLFGroupShape) {
            XSLFGroupShape group = (XSLFGroupShape) shape;
            for (XSLFShape subShape : group.getShapes()) {
                dealGroupShape(subShape);
            }
        } else {
            System.out.println("layout:   " + shape.getClass());
        }
    }
}
