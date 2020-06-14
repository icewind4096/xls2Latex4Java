package com.mindmotion.xls2latex.convert;

import com.mindmotion.xls2latex.domain.ParamaterInfo;
import com.mindmotion.xls2latex.enums.GeneralFileTypeEnum;
import com.mindmotion.xls2latex.enums.HAligmentEnum;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Arg2Paramater {
    public static ParamaterInfo arg2Paramater(String[] args){
        ParamaterInfo paramaterInfo = new ParamaterInfo();
        paramaterInfo.setDestFileName(args[0]);
        paramaterInfo.setSourceFileName(args[1]);
        paramaterInfo.setGeneralFileTypeEnum(GeneralFileTypeEnum.values()[Integer.parseInt(args[2])]);
        paramaterInfo.setWidth(Integer.parseInt(args[3]));
        paramaterInfo.setColCount(Integer.parseInt(args[4]));
        paramaterInfo.setColWidths(arg2ColWidths(args[5]));
        paramaterInfo.setLanguage(Integer.parseInt(args[6]));
        paramaterInfo.setAligment(HAligmentEnum.values()[Integer.parseInt(args[7])]);
        if (paramaterInfo.getGeneralFileTypeEnum() != GeneralFileTypeEnum.REGFILE){
            paramaterInfo.setTableName(args[8]);
        }
        return paramaterInfo;
    }

    private static List<Integer> arg2ColWidths(String colWidths) {
        List<String> strList = Arrays.asList(colWidths.split(","));
        List<Integer> intList = new ArrayList<Integer>();
        for (String str : strList) {
            intList.add(Integer.parseInt(str));
        }
        return intList;
    }
}
