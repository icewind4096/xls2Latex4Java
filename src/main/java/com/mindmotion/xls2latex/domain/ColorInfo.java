package com.mindmotion.xls2latex.domain;

public class ColorInfo {
    private int r;
    private int g;
    private int b;

    public ColorInfo(int r, int g, int b) {
        this.r = r;
        this.g = g;
        this.b = b;
    }

    public Integer toRGB(){
        return r << 16 | g << 8 | b;
    }
}
