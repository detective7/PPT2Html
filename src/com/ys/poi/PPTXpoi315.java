package com.ys.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.ys.util.Wmf2Svg;

public class PPTXpoi315 {

    // 图片存放路径
    public final static String path = "E:\\PPTpoi\\y\\imgPPTX\\";
    private static Map img;

    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("E:\\PPTpoi\\y\\unit 4 My home.pptx"));
        img = new HashMap();

        for (XSLFPictureData pict : ppt.getPictureData()) {
            byte[] data = pict.getData();
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

    }

}
