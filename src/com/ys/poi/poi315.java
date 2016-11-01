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

import com.ys.util.InsertUtil;
import com.ys.util.TransformUtil;
import com.ys.util.Wmf2Svg;

public class poi315 {

    // 图片默认存放路径
    public final static String path = "E:\\PPTpoi\\s\\";
    // public final static String pathS = "E:\\PPTpoi\\安全用电\\sound\\";
    private static Map<Integer, String> img;
    private static List<PPTText> texts;
    private static PrintStream printStream;
    private static FileOutputStream fs;
    private static InsertUtil insertUtil;

    public static void main(String[] args) throws Exception {
        /*
         * 网页输出
         */
        fs = new FileOutputStream(new File(path + "output.html"));
        printStream = new PrintStream(fs);

        // 加载PPT
        HSLFSlideShow ss = new HSLFSlideShow(new HSLFSlideShowImpl(path + "s.ppt"));
        img = new HashMap<Integer, String>();

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
            String ext = pict.getType().extension;
            FileOutputStream out = new FileOutputStream(path + "img\\" + pict.getIndex() + ext);
            out.write(data);
            if (ext.equals(".wmf")) {
                Wmf2Svg.convert(path + "img\\" + pict.getIndex() + ext);
                ext = ".svg";
            }
            // System.out.println(pict.getHeader().toString());
            img.put(pict.getIndex(), pict.getIndex() + ext);
            out.close();
        }

        // 存字相关信息
        texts = new ArrayList<PPTText>();
        List<HSLFSlide> slides = ss.getSlides();

        // 获取所有不同字体的文字
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

            // 背景（母版，主题，布局里面的取不到）
            /*if (slide.getBackground().getFill().getPictureData() != null) {
                System.out.println(slide.getSlideNumber() + ": " + slide.getBackground().getFill().getPictureData().getIndex());
            } else {
                System.out.println(slide.getSlideNumber() + ": " + slide.getBackground().getFill().getBackgroundColor());
            }*/
            
            //布局的动画
            /*org.apache.poi.hslf.record.Slide sRecord = slide.getSlideRecord();
            AnimationInfo animInfo = (AnimationInfo) sRecord.findFirstOfType(RecordTypes.AnimationInfo.typeID);
            if (animInfo!=null){ 
                System.out.println(slide.getSlideNumber()+"  "+animInfo.getAnimationInfoAtom().toString());
            }else{
                System.out.println(slide.getSlideNumber()+"  "+"no animInfo");
            }*/
        }

        insertUtil = new InsertUtil(printStream, texts, img);

        insertUtil.JSswitch(slides.size());

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
            //插入背景
            if (slides.get(i).getBackground().getFill().getPictureData() != null) {
                insertUtil.insertImg(i, slides.get(i).getBackground().getFill().getPictureData().getIndex(), "[x=0,y=0,w="+ss.getPageSize().width+",h="+ss.getPageSize().height+"]");
            } /*else {
                System.out.println(slides.get(i).getSlideNumber() + ": " + slides.get(i).getBackground().getFill().getBackgroundColor());
            }*/
            int j = 0;
            for (HSLFShape shape : slides.get(i).getShapes()) {
                /*
                 * 无效动画
                 * EscherContainerRecord container = shape.getSpContainer();
                ArrayList lAnimInfoAtom = new ArrayList();
                container.getRecordsById((short) RecordTypes.AnimationInfoAtom.typeID, lAnimInfoAtom);
                if (lAnimInfoAtom.size() != 0)
                {
                  // unknown should be of type AnimationInfoAtom...
                  UnknownEscherRecord unknown = (UnknownEscherRecord)lAnimInfoAtom.get(0);
                  System.out.println(unknown.getRecordName());
                } */
                // System.out.println("框类型" + shape.getClass().toGenericString()
                // + " " + shape.getShapeName() + " ");
                insertUtil.dealWithGroup(j, slides.get(i).getSlideNumber(), shape);
                j++;
            }
            // System.out.println("第" + slides.get(i).getSlideNumber() +
            // "张PPT解析结束 \n");
            printStream.println("}}");
        }
        insertUtil.endDiv(ss.getPageSize().width, ss.getPageSize().height);
        printStream.close();
        ss.close();
    }

}
