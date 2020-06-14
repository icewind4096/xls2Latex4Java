package com.mindmotion.xls2latex.domain;

import com.mindmotion.xls2latex.enums.GeneralFileTypeEnum;
import com.mindmotion.xls2latex.enums.HAligmentEnum;

import java.util.ArrayList;
import java.util.List;

public class ParamaterInfo {
    String destFileName;
    String sourceFileName;
    String tableName;
    GeneralFileTypeEnum generalFileTypeEnum;
    Integer width;
    Integer colCount;
    Integer language;
    HAligmentEnum aligment;
    List<Integer> colWidths;

    public ParamaterInfo() {
        colWidths = new ArrayList<Integer>();
    }

    public List<Integer> getColWidths() {
        return colWidths;
    }

    public void setColWidths(List<Integer> colWidths) {
        this.colWidths = colWidths;
    }

    public String getDestFileName() {
        return destFileName;
    }

    public void setDestFileName(String destFileName) {
        this.destFileName = destFileName;
    }

    public String getSourceFileName() {
        return sourceFileName;
    }

    public void setSourceFileName(String sourceFileName) {
        this.sourceFileName = sourceFileName;
    }

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public GeneralFileTypeEnum getGeneralFileTypeEnum() {
        return generalFileTypeEnum;
    }

    public void setGeneralFileTypeEnum(GeneralFileTypeEnum generalFileTypeEnum) {
        this.generalFileTypeEnum = generalFileTypeEnum;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }

    public Integer getColCount() {
        return colCount;
    }

    public void setColCount(Integer colCount) {
        this.colCount = colCount;
    }

    public Integer getLanguage() {
        return language;
    }

    public void setLanguage(Integer language) {
        this.language = language;
    }

    public HAligmentEnum getAligment() {
        return aligment;
    }

    public void setAligment(HAligmentEnum aligment) {
        this.aligment = aligment;
    }
}
