package com.mindmotion.xls2latex.util;

import org.apache.commons.lang3.StringUtils;

public class LatexUtil {
    public static String conver2LatexString(String text, boolean usePar){
        text = StringUtils.replace(text, "_", "\\_");
        if (usePar == true){
            return StringUtils.replace(text, "\n", " \\par ");
        } else {
            return StringUtils.replace(text, "\n", " \\\\ ");
        }
    }

    public static String getRegDisplayText(String value, int length) {
        String text = LatexUtil.conver2LatexString(value, true);
        return StringUtils.rightPad(text, length, " ");
    }
}
