package com.mindmotion.xls2latex.enums;

public enum ResultEnum {
    SUCCESS(0, "SUCCESS"),
    READEXCELFILEFAIL(1, "Read Excel File Fail"),
    MAKEOUTDIRFAIL(2, "Make Out Directory Fail"),
    EXCELFILENOTEXIST(3, "Excel Not Exist"),
    ;

    private int code;
    private String descript;

    public int getCode() {
        return code;
    }

    public String getDescript() {
        return descript;
    }

    ResultEnum(int code, String descript) {
        this.code = code;
        this.descript = descript;
    }
}
