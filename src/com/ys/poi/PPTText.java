package com.ys.poi;

public class PPTText {
    private int slideNum;
    private String text;
    private Double size;
    private String fontFamily;
    private String color;

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public Double getSize() {
        return size;
    }

    public void setSize(Double size) {
        this.size = size;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public int getSlideNum() {
        return slideNum;
    }

    public void setSlideNum(int slideNum) {
        this.slideNum = slideNum;
    }

    @Override
    public String toString() {
        return "PPTText [slideNum=" + slideNum + ", text=" + text + ", size=" + size + ", fontFamily=" + fontFamily + ", color=" + color + "]\n";
    }
    
}
