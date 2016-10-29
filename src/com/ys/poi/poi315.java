package com.ys.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.sl.usermodel.PictureData;

import com.ys.util.InsertUtil;
import com.ys.util.TransformUtil;
import com.ys.util.Wmf2Svg;

public class poi315 {

    // 图片默认存放路径
    public final static String path = "E:\\PPTpoi\\y\\img\\";
    // public final static String pathS = "E:\\PPTpoi\\安全用电\\sound\\";
    private static Map img;
    private static List<PPTText> texts;
    private static PrintStream printStream;
    private static FileOutputStream fs;
    private static InsertUtil insertUtil;

    public static void main(String[] args) throws Exception {
        /*
         * 网页输出
         */
        fs = new FileOutputStream(new File("E:\\PPTpoi\\y\\output.html"));
        printStream = new PrintStream(fs);
        printStream.println(
                "<!DOCTYPE html>\n<html lang=\"en\" xmlns=\"http://www.w3.org/1/xhtml\">\n<head>\n<meta charset=\"utf-8\" />" + "<title>js实现ppt</title>");

        // 加载PPT
        HSLFSlideShow ss = new HSLFSlideShow(new HSLFSlideShowImpl("E:\\PPTpoi\\y\\y.ppt"));
        img = new HashMap();

        // HSLFSoundData[] sds = ss.getSoundData();
        // for(HSLFSoundData sd:ss.getSoundData()){
        // System.out.println(sds.length);
        // byte[] sData = sd.getData();
        // String sType= sd.getSoundType();
        // String sName=sd.getSoundName();
        // FileOutputStream out = new FileOutputStream(pathS + sName);
        // out.write(sData);
        // out.close();
        // }

        // 获取所有图片，并把矢量图wmf转为网页可显示的svg格式。 extract all pictures contained in the
        // presentation
        for (HSLFPictureData pict : ss.getPictureData()) {
            // picture data
            byte[] data = pict.getData();
            // System.out.println(pict.getHeader().toString());
            PictureData.PictureType type = pict.getType();
            String ext = type.extension;
            FileOutputStream out = new FileOutputStream(path + pict.getIndex() + ext);
            out.write(data);
            if (ext.equals(".wmf")) {
                Wmf2Svg.convert(path + pict.getIndex() + ext);
                ext = ".svg";
            }
            // System.out.println(pict.getHeader().toString());
            img.put(pict.getIndex(), pict.getIndex() + ext);
            out.close();
        }

        // 存字相关信息
        texts = new ArrayList<PPTText>();
        List<HSLFSlide> slides = ss.getSlides();

        //获取所有不同字体的文字
        for (HSLFSlide slide : slides) {
            // 取每部分字的字体，大小和颜色
            List<List<HSLFTextParagraph>> textPss = slide.getTextParagraphs();
            for (List<HSLFTextParagraph> textPs : textPss) {
                for (HSLFTextParagraph textP : textPs) {
                    List<HSLFTextRun> trs = textP.getTextRuns();
                    for (HSLFTextRun tr : trs) {
                        PPTText text = new PPTText();
                        String t = tr.getRawText().replaceAll("", "");
                        if (t != null && !t.equals("") && !t.trim().equals("")) {
                            text.setSlideNum(slide.getSlideNumber());
                            text.setText(t.replaceAll("[\n\r\t]", ""));
                            text.setColor(TransformUtil.toHex(tr.getFontColor().getSolidColor().getColor().toString()));
                            text.setSize(tr.getFontSize());
                            text.setFontFamily(tr.getFontFamily());
                            texts.add(text);
                        }
                    }
                }
            }
            
            insertUtil=new InsertUtil(printStream,texts,img);

            printStream.println(insertUtil.JSswitch(slides.size()));

            // 下面备注
            // System.out.println(slide.getSlideNumber() + ": " +
            // slide.getNotes().getTextParagraphs().toString());
            // if (slide.getBackground().getFill().getPictureData() != null) {
            // System.out.println(slide.getSlideNumber() + ": " +
            // slide.getBackground().getFill().getPictureData().getIndex());
            // } else {
            // System.out.println(slide.getSlideNumber() + ": " +
            // slide.getBackground().getFill().getBackgroundColor());
            // }
            // System.out.println(slide.getSlideNumber() + ": " +
            // slide.getBackground().getFill().getForegroundColor());
            // System.out.println(slide.getSlideNumber() + ": " +
            // slide.getFollowMasterBackground());
            // System.out.println(slide.getSlideNumber() + ": " +
            // slide.getMasterSheet().getColorScheme().getColor(2));
        }

        // 获取母版
        // for (HSLFSlideMaster am : ss.getSlideMasters()) {
        // System.out.println(am.toString());
        // if (am.getBackground().getFill().getPictureData() != null) {
        // System.out.println(ss.getSlideMasters().size()+" "+am.toString()+"
        // "+am.getBackground().getFill().getPictureData().getIndex());
        // }
        // }

        // System.out.println(texts);
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
                insertUtil.dealWithGroup(j, slides.get(i).getSlideNumber(), shape);
                j++;
            }
            // System.out.println("第" + slides.get(i).getSlideNumber() +
            // "张PPT解析结束 \n");
            printStream.println("}}");
        }
        printStream.println(insertUtil.endDiv(ss.getPageSize().width, ss.getPageSize().height));
        printStream.close();
    }

}
