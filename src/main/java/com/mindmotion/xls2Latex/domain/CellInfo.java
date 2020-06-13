package com.mindmotion.xls2Latex.domain;

import com.mindmotion.xls2Latex.common.Rect;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class CellInfo {
    private String text;
    private HorizontalAlignment hAligment;
    private VerticalAlignment vAligment;
    private Boolean merged;
    private Rect rect;
    private Integer colWidth;
    private Integer fontColor;
    private Integer backColor;

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public HorizontalAlignment gethAligment() {
        return hAligment;
    }

    public void sethAligment(HorizontalAlignment hAligment) {
        this.hAligment = hAligment;
    }

    public VerticalAlignment getvAligment() {
        return vAligment;
    }

    public void setvAligment(VerticalAlignment vAligment) {
        this.vAligment = vAligment;
    }

    public Boolean getMerged() {
        return merged;
    }

    public void setMerged(Boolean merged) {
        this.merged = merged;
    }

    public Rect getRect() {
        return rect;
    }

    public void setRect(Rect rect) {
        this.rect = rect;
    }

    public Integer getColWidth() {
        return colWidth;
    }

    public void setColWidth(Integer colWidth) {
        this.colWidth = colWidth;
    }

    public Integer getFontColor() {
        return fontColor;
    }

    public void setFontColor(Integer fontColor) {
        this.fontColor = fontColor;
    }

    public Integer getBackColor() {
        return backColor;
    }

    public void setBackColor(Integer backColor) {
        this.backColor = backColor;
    }
}
