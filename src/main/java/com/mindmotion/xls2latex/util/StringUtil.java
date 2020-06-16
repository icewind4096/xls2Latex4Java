package com.mindmotion.xls2latex.util;

public class StringUtil {
    public static String colorToRGB(int color){
        int R, G, B;
        R = color & 0xFF;
        G = (color & 0xFF00) >> 8;
        B = (color & 0xFF0000) >> 16;

        return String.format("%d,%d,%d", R, G, B);
    }
}
