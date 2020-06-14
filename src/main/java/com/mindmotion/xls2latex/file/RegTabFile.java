package com.mindmotion.xls2latex.file;

import com.mindmotion.xls2latex.domain.CellInfo;
import com.mindmotion.xls2latex.domain.ParamaterInfo;
import com.mindmotion.xls2latex.enums.ResultEnum;
import com.mindmotion.xls2latex.util.FileUtil;
import com.mindmotion.xls2latex.util.LatexUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class RegTabFile {
    public static Integer GenerateFile(ParamaterInfo paramaterInfo){
        List<List<CellInfo>> rowDatas = new ArrayList<List<CellInfo>>();
        List<String> latexDatas = new ArrayList<String>();
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(paramaterInfo.getSourceFileName()));
            for (int i = 0; i < workbook.getNumberOfSheets(); i ++){
                clearRowDatas(rowDatas);
                clearLatexDatas(latexDatas);
                if (ExcelFile.readSheetData(workbook.getSheetAt(i), rowDatas, 0, -1, paramaterInfo.getColCount())){
                    translate2RegTab(paramaterInfo.getLanguage(), rowDatas, latexDatas);
                    if (!generateRegFile(getRegTabFileName(paramaterInfo.getDestFileName(), getRegTabFileNamePrefix(paramaterInfo.getSourceFileName()), workbook.getSheetAt(i).getSheetName()), latexDatas)){
                        return ResultEnum.MAKEOUTFILEFAIL.getCode();
                    }
                } else {
                    return ResultEnum.READEXCELFILEFAIL.getCode();
                }
            }
            return ResultEnum.SUCCESS.getCode();
        } catch (IOException e) {
            e.printStackTrace();
            return ResultEnum.READEXCELFILEFAIL.getCode();
        } finally {
            try {
                if (workbook != null){
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static boolean generateRegFile(String fileName, List<String> datas) {
        return FileUtil.saveToFileByList(fileName, datas);
    }

    private static void clearLatexDatas(List<String> datas) {
        datas.clear();
    }

    private static String getRegTabFileNamePrefix(String fileName) {
        if (fileName.indexOf(".") == -1) {
            return fileName;
        };

        String value = StringUtils.substringBeforeLast(fileName, ".");
        if (value.indexOf("\\") == -1){
            return value;
        }

        return StringUtils.substringAfterLast(value, "\\");
    }

    private static String getRegTabFileName(String directory, String prefix, String sheetName) {
        return String.format("%s\\%s_%s.tex", directory, prefix, sheetName);
    }

    private static boolean translate2RegTab(int language, List<List<CellInfo>> sourceDatas, List<String> targetDatas) {
        generalRegHead(targetDatas, language);

        generalRegBody(targetDatas, sourceDatas);

        return true;
    }

    private static void generalRegBody(List<String> lists, List<List<CellInfo>> datas) {
        lists.add("{");
        for (int i = 0; i < datas.size(); i ++){
            lists.add(dataToRegLatex(datas.get(i)));
        }
        lists.add("}");
    }

    private static String dataToRegLatex(List<CellInfo> cellInfos) {
        String text = String.format("%s &%s &%s &%s ", getRegBitPosDisplay(cellInfos.get(0).getText())
                                                     , getRegBitNameDisplay(cellInfos.get(1).getText())
                                                     , getRegBitAttrDisplay(cellInfos.get(2).getText())
                                                     , getRegBitRestValueDisplay(cellInfos.get(3).getText()));

        int textLength = text.length();
        String[] regBitDescripts = getRegBitDescriptDisplay(cellInfos.get(4).getText()).split("\\\\par");
        for (int i = 0; i < regBitDescripts.length; i++) {
            if (i == 0) {
                text = text + String.format(" &%s", regBitDescripts[i]);
            } else {
                text = text + StringUtils.leftPad(regBitDescripts[i], textLength + regBitDescripts[i].length(), " ");
            }
            if (i < regBitDescripts.length - 1){
                text = text + "\\par" + "\r\n";
            }
        }

        return text + "\\\\";
    }

    private static String getRegBitDescriptDisplay(String value) {
        return LatexUtil.getRegDisplayText(value, 0);
    }

    private static String getRegBitRestValueDisplay(String value) {
        return LatexUtil.getRegDisplayText(value, 10);
    }

    private static String getRegBitAttrDisplay(String value) {
        return LatexUtil.getRegDisplayText(value, 6);
    }

    private static String getRegBitNameDisplay(String value) {
        return LatexUtil.getRegDisplayText(value, 16);
    }

    private static String getRegBitPosDisplay(String value) {
        return LatexUtil.getRegDisplayText(value, 12);
    }

    private static void generalRegHead(List<String> lists, int language) {
        if (language == 0) {
            lists.add("regDescriptionCN");
        }
        else {
            lists.add("regDescriptionEN");
        }
    }

    private static void clearRowDatas(List<List<CellInfo>> rowData) {
        for (int i = 0; i < rowData.size() - 1; i++ ){
            clearRowData(rowData.get(i));
        }
        rowData.clear();
    }

    private static void clearRowData(List<CellInfo> rowData) {
        rowData.clear();
    }
}
