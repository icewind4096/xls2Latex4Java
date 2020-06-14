package com.mindmotion.xls2latex.enums;

public enum GeneralFileTypeEnum {
    REGOVERVIEWFILE(0, "REG OVERVIEW FILE"),
    REGFILE(1, "REG FILE"),
    GENERALFILE(2, "GENERAL FILE"),
    ;

    private int code;
    private String descript;

    public int getCode() {
        return code;
    }

    public String getDescript() {
        return descript;
    }

    GeneralFileTypeEnum(int code, String descript) {
        this.code = code;
        this.descript = descript;
    }
}
