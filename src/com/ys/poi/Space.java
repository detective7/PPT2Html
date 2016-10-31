package com.ys.poi;

import java.awt.Dimension;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PrintStream;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class Space {

    private static PrintStream out;
    private static String path;
    
    public static void main(String[] args) throws Exception{
        path="E:\\PPTpoi\\y\\";
        out = System.out;

        FileInputStream is = new FileInputStream(path+"unit 4 My home.pptx");
        XMLSlideShow ppt = new XMLSlideShow(is);
        is.close();

        getTextPic(ppt);

        ppt.close();
    }
    
    //获取文字图片
    public static void getTextPic(XMLSlideShow ppt) throws Exception{
        for (PackagePart p : ppt.getAllEmbedds()) {
            String type = p.getContentType();

            String name = p.getPartName().getName();
            out.println("Embedded file (" + type + "): " + name);

            InputStream pIs = p.getInputStream();

            pIs.close();
        }

        for (XSLFPictureData data : ppt.getPictureData()) {
            String type = data.getContentType();
            String name = data.getFileName();
            out.println("Picture (" + type + "): " + name);

            FileOutputStream fileout = new FileOutputStream(path+"imgPPTX\\" + name);
            fileout.write(data.getData());
            fileout.close();
            InputStream pIs = data.getInputStream();
            pIs.close();
        }

        Dimension pageSize = ppt.getPageSize();
        out.println("Pagesize: " + pageSize);

        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape shape : slide) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape txShape = (XSLFTextShape) shape;
                    String txStr = txShape.getText().replaceAll("\n", "<br>");
                    out.println("txStr:   =>"+txStr);
                    for (XSLFTextParagraph p : txShape) {
//                        out.println("Paragraph level: " + p.getIndentLevel());
                        for (XSLFTextRun r : p) {
                            String reStr = r.getRawText().replaceAll("\n", "<br>");
                            System.out.println("deStr:   ->"+reStr);
                            if (reStr.equals("[?]") || reStr.startsWith("?")) {
                                txStr = txStr.replaceFirst("\\" + reStr, "");
                            } else {
                                txStr=txStr.replaceFirst(reStr, "");
                            }
                            for (; txStr.startsWith(" ");) {
                                txStr = txStr.replaceFirst(" ", "");
                            }
                            // 回车换行得额外弄，按顺序把字符替换掉,剩下的，是上面textparam查不出的字符，再贴上去
                            for (; txStr.startsWith("<br>");) {
                                reStr = reStr + "<br>";
                                txStr = txStr.replaceFirst("<br>", "");
                            }
                            out.println("txStr:   =>"+txStr);
//                            System.out.println("  bold: " + r.isBold());
//                            System.out.println("  italic: " + r.isItalic());
//                            System.out.println("  underline: " + r.isUnderlined());
//                            System.out.println("  font.family: " + r.getFontFamily());
//                            System.out.println("  font.size: " + r.getFontSize());
//                            System.out.println("  font.color: " + r.getFontColor());
                        }
                    }
                } else if (shape instanceof XSLFPictureShape) {
                    XSLFPictureShape pShape = (XSLFPictureShape) shape;
                    XSLFPictureData pData = pShape.getPictureData();
                    out.println(pData.getFileName());
                } else {
//                    out.println("Process me: " + shape.getClass());
                }
            }
        }
    }
}
