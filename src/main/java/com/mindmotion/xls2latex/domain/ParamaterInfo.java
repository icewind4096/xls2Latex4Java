package com.mindmotion.xls2latex.domain;

import java.util.ArrayList;
import java.util.List;

public class ParamaterInfo {
    String destDirectory;
    String sourceFileName;
    String tableName;
    Integer type;
    Integer width;
    Integer colCount;
    Integer language;
    Integer aligment;
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

    public String getDestDirectory() {
        return destDirectory;
    }

    public void setDestDirectory(String destDirectory) {
        this.destDirectory = destDirectory;
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

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
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

    public Integer getAligment() {
        return aligment;
    }

    public void setAligment(Integer aligment) {
        this.aligment = aligment;
    }
}
