package com.mindmotion.xls2latex.file;

import com.mindmotion.xls2latex.domain.CellInfo;
import com.mindmotion.xls2latex.domain.ParamaterInfo;
import com.mindmotion.xls2latex.enums.GeneralFileTypeEnum;
import com.mindmotion.xls2latex.enums.HAligmentEnum;
import com.mindmotion.xls2latex.enums.ResultEnum;
import com.mindmotion.xls2latex.util.FileUtil;
import com.mindmotion.xls2latex.util.LatexUtil;
import com.mindmotion.xls2latex.util.StringUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class GeneralTabFile {
    public static int GenerateFile(ParamaterInfo paramaterInfo) {
        List<CellInfo> columnDatas = new ArrayList<CellInfo>();
        List<List<CellInfo>> rowDatas = new ArrayList<List<CellInfo>>();
        List<String> latexDatas = new ArrayList<String>();
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(paramaterInfo.getSourceFileName()));
            if (readSheetColumnData(workbook.getSheetAt(0), columnDatas, paramaterInfo.getColCount(), paramaterInfo.getColWidths())
             && readSheetRowData(workbook.getSheetAt(0), rowDatas, paramaterInfo.getColCount(), paramaterInfo.getColWidths())){
                translate2GeneralTab(paramaterInfo.getGeneralFileTypeEnum(), paramaterInfo.getTableName(), paramaterInfo.getTableName(), paramaterInfo.getColCount(), paramaterInfo.getLanguage(), paramaterInfo.getAligment(),  columnDatas, rowDatas, latexDatas);
                if (!generateGeneralFile(paramaterInfo.getDestFileName(), latexDatas)){
                    return ResultEnum.MAKEOUTFILEFAIL.getCode();
                }
            } else {
                return ResultEnum.READEXCELFILEFAIL.getCode();
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

    private static boolean generateGeneralFile(String fileName, List<String> datas) {
        return FileUtil.saveToFileByList(fileName, datas);
    }

    private static boolean translate2GeneralTab(GeneralFileTypeEnum generalFileType, String tableName, String tableLabel, Integer colCount, Integer language, HAligmentEnum aligment, List<CellInfo> columnDatas, List<List<CellInfo>> rowDatas, List<String> targetDatas) {
        targetDatas.add("\\captionsetup[table]{");
        //label左对齐
        targetDatas.add("	singlelinecheck=false,");
        targetDatas.add("}");

        //设置表格label与上一段文字的距离
        targetDatas.add("\\setlength\\intextsep{10pt}");
        targetDatas.add("\\makeatletter");
        targetDatas.add("\\newcommand{\\hlineNpx}[1]{\\noalign {\\ifnum 0=`}\\fi \\hrule height #1pt	\\futurelet \\reserved@a \\@xhline}");
        targetDatas.add(getFontSetByLanguage(language));

        //设置行间距
        targetDatas.add("\\linespread{1.2}");

        //表格开始
        targetDatas.add("\\begin{table}[h!]");
        //表格和label之间的距离
        targetDatas.add("\\setlength{\\abovecaptionskip}{1pt}");
        //表格和label之间的距离
        targetDatas.add("\\setlength{\\belowcaptionskip}{10pt}");
        //表格对齐位置开始
        targetDatas.add(getTableAligmentBegin(aligment));
        //表格标题
        targetDatas.add(String.format("\\caption{%s}", tableName));
        //表格引用标号
        targetDatas.add(String.format("\\label{tab:%s}", tableLabel));
        //表格列头对齐方式开始
        targetDatas.add(String.format("\\begin{tabular}{%s}", getColumnStyle(columnDatas, generalFileType)));
        targetDatas.add("\\hlineNpx{1}");

        //表格列头数据
        targetDatas.add(getColumnData(columnDatas, generalFileType));
        targetDatas.add("\\hlineNpx{1}");

        //得到表格数据
        targetDatas.add(getBodyData(rowDatas, generalFileType));
        targetDatas.add("\\hlineNpx{1}");
        //表格列头对齐方式结束
        targetDatas.add("\\end{tabular}");
        //表格对齐位置结束
        targetDatas.add(getTableAligmentEnd(aligment));

        //表格结束
        targetDatas.add("\\end{table}");

        return true;
    }

    private static String getBodyData(List<List<CellInfo>> rowDatas, GeneralFileTypeEnum generalFileType) {
        String text = "";
        for (int i = 0; i < rowDatas.size(); i ++){
            text = text + getRowData(i, false, rowDatas.get(i), generalFileType) + "\r\n";
        }
        if (text.length() > 0){
            text = text.substring(0, text.length() - 2);
        }
        return text;
    }

    private static String getColumnData(List<CellInfo> columnDatas, GeneralFileTypeEnum generalFileType) {
        return getRowData(0, true, columnDatas, generalFileType);
    }

    private static String getRowData(int rowIndex, boolean isColumn, List<CellInfo> datas, GeneralFileTypeEnum generalFileType) {
        CellInfo cellInfo = null;
        String text = "";
        for (int i = 0; i < datas.size(); i++){
            cellInfo = datas.get(i);
            if (cellInfo.getMerged() == true){
                if (!cellInfo.getRect().getBottom().equals(cellInfo.getRect().getTop())){
                    if (rowIndex + 1 == cellInfo.getRect().getTop()){
                        text = text + String.format("\\multirow{%d}{%dpt}{%s}", cellInfo.getRect().getBottom() - cellInfo.getRect().getTop() + 1, cellInfo.getColWidth(), getCellData(cellInfo.gethAligment(), isColumn, cellInfo.getColWidth(), cellInfo.getText(), cellInfo.getBackColor(), cellInfo.getFontColor(), generalFileType, isAutoRef(generalFileType, i, datas.size() - 1)));
                    }
                } else {
                    if (!cellInfo.getRect().getLeft().equals(cellInfo.getRect().getRight())){
                        if (i == cellInfo.getRect().getLeft()){
                            text = text + String.format("\\multicolumn{%d}{%s}{%s} " , cellInfo.getRect().getRight() - cellInfo.getRect().getLeft() + 1
                                                                                     , getCellVLineString(cellInfo.gethAligment(), getCellLeftLineVisable(datas, i), getCellRightLineVisiable(datas, i))
                                                                                     , getCellData(cellInfo.gethAligment(), isColumn, cellInfo.getColWidth(), cellInfo.getText(), cellInfo.getBackColor(), cellInfo.getFontColor(), generalFileType, isAutoRef(generalFileType, i, datas.size() - 1)));
                        } else {
                            continue;
                        }
                    }
                }
            } else {
                text = text + String.format("%s", getCellData(cellInfo.gethAligment(), isColumn, cellInfo.getColWidth(), cellInfo.getText(), cellInfo.getBackColor(), cellInfo.getFontColor(), generalFileType, isAutoRef(generalFileType, i, datas.size() - 1)));
            }
            text = text + " &";
        }
        text = text.substring(0, text.length() - 1);

        return text + String.format("\\\\ %s ", getCLine(rowIndex, isColumn, datas));
    }

    private static String getCellVLineString(HorizontalAlignment horizontalAlignment, boolean cellLeftLineVisable, boolean cellRightLineVisiable) {
        String text = cellHAligment2String(horizontalAlignment);

        if (cellLeftLineVisable){
            text = "|" + text;
        }

        if (cellRightLineVisiable){
            text = text + "|";
        }

        return text;
    }

    private static String cellHAligment2String(HorizontalAlignment horizontalAlignment) {
        switch (horizontalAlignment){
            case CENTER: return "c";
            case RIGHT: return "r";
            default: return "l";
        }
    }

    private static boolean getCellLeftLineVisable(List<CellInfo> datas, int colIndex) {
        if (colIndex == 0){
            return false;
        } else {
            return !datas.get(colIndex - 1).getRect().getRight().equals(datas.get(colIndex - 1).getRect().getLeft());
        }
    }

    private static boolean getCellRightLineVisiable(List<CellInfo> datas, int colIndex) {
        if (datas.get(colIndex).getRect().getRight() == datas.size() - 1){
            return false;
        } else {
            return !datas.get(datas.get(colIndex).getRect().getRight()).getMerged();
        }
    }

    private static String getCLine(int rowIndex, boolean isColumn, List<CellInfo> datas) {
        String text = "";
        for (int i = 0; i < datas.size(); i ++){
            CellInfo cellInfo = datas.get(i);
            if ((cellInfo.getMerged() == false) || (isColumn == true) || (isColumn == false && cellInfo.getMerged() == true && cellInfo.getRect().getBottom() == rowIndex + 1)){
                text = text + String.format(" \\cline{%d-%d}", i + 1, i + 1);
            }
        }
        return text;
    }

    private static String getCellData(HorizontalAlignment horizontalAlignment, boolean isColumn, Integer cellWidth, String text, Integer backColor, Integer fontColor, GeneralFileTypeEnum generalFileType, boolean autoRef) {
        String value;
        if (isColumn){
            if (generalFileType == GeneralFileTypeEnum.REGOVERVIEWFILE){
                value = String.format("\\columnTitle\\cell\\cellcolor[RGB]{%s}{%s}", StringUtil.colorToRGB(0xB0D5FA), text);
            } else {
                value = String.format("\\columnTitle{%s}", text);
            }
        } else {
            value = String.format("\\cell{%s}", text);
        }

        value = String.format("\\begin{minipage}[]{%dpt}\\begin{%s}\\vspace{5pt} %s \\vspace{5pt}\\end{%s} \\end{minipage}", cellWidth, HAligment2FlushString(horizontalAlignment), value, HAligment2FlushString(horizontalAlignment));

        if (!autoRef){
            return LatexUtil.conver2LatexString(value, false);
        } else {
            return value;
        }
    }

    private static String HAligment2FlushString(HorizontalAlignment horizontalAlignment) {
        switch (horizontalAlignment){
            case CENTER: return "center";
            case LEFT: return "flushleft";
            default: return "flushright";
        }
    }

    private static boolean isAutoRef(GeneralFileTypeEnum generalFileType, int colIndex, int colCount) {
        if (generalFileType == GeneralFileTypeEnum.REGOVERVIEWFILE) {
            return colIndex == colCount;
        }
        return false;
    }

    private static String getColumnStyle(List<CellInfo> columnDatas, GeneralFileTypeEnum generalFileType) {
        String text = "";
        for (CellInfo cellInfo : columnDatas) {
            text = text + String.format("%sp{%dpt}<{\\%s}", getColumnVLineVisiable(generalFileType), cellInfo.getColWidth(), columnHAlignment2String(cellInfo.gethAligment()));
        }

        if (generalFileType == GeneralFileTypeEnum.GENERALFILE){
            text = text.substring(1);
        }

        return text;
    }

    private static String columnHAlignment2String(HorizontalAlignment aligment) {
        //左右是不是写反了??
        switch (aligment){
            case RIGHT: return "raggedleft";
            case CENTER: return "centering";
            default: return "raggedright";
        }
    }

    private static String getColumnVLineVisiable(GeneralFileTypeEnum generalFileType) {
        if (generalFileType == GeneralFileTypeEnum.GENERALFILE){
            return "|";
        } else {
            return "";
        }
    }

    private static String getTableAligmentEnd(HAligmentEnum aligment) {
        switch (aligment){
            case center: return "\\end{center}";
            case right: return "\\end{flushright}";
            default: return "\\end{flushleft}";
        }
    }

    private static String getTableAligmentBegin(HAligmentEnum aligment) {
        switch (aligment){
            case center: return "\\begin{center}";
            case right: return "\\begin{flushright}";
            default: return "\\begin{flushleft}";
        }
    }

    private static String getFontSetByLanguage(Integer language) {
        if (language == 0){
            return getFontSetByCN();
        } else {
            return getFontSetByEN();
        }
    }

    private static String getFontSetByEN() {
        return getFontSet("宋体", "9", "RobotoMono", "9");
    }

    private static String getFontSetByCN() {
        return getFontSet("宋体", "9", "RobotoMono", "9");
    }

    private static String getFontSet(String columnForntName, String columnFontSize, String cellFontName, String cellFontSize) {
        return String.format( "\\setCJKfamilyfont{ltzt}{%s}\r\n"
                            + "\\newcommand{\\dyzt}{\\%s}\r\n"
                            + "\\newcommand{\\ltzt}{\\CJKfamily{ltzt}}\r\n"
                            + "\\newcommand{\\ltzh}{\\fontsize{%spt}{\\baselineskip}\\selectfont}\r\n"
                            + "\\newcommand{\\dyzh}{\\fontsize{%spt}{\\baselineskip}\\selectfont}\r\n"
                            + "\\newcommand{\\columnTitle}{\\ltzt \\ltzh \\bfseries}\r\n"
                            + "\\newcommand{\\cell}{\\dyzt \\dyzh}\r\n", columnForntName, cellFontName, columnFontSize, cellFontSize);
    }

    private static boolean readSheetRowData(Sheet sheet, List<List<CellInfo>> rowDatas, Integer colCount, List<Integer> colWidths) {
        if (ExcelFile.readSheetData(sheet, rowDatas, 1, -1, colCount) == true){
            for (int i = 0; i < rowDatas.size(); i ++){
                fillColWidthData(rowDatas.get(i), colWidths);
            }
            return true;
        }
        return false;
    }

    private static boolean readSheetColumnData(Sheet sheet, List<CellInfo> columnDatas, int colCount, List<Integer> colWidths) {
        List<List<CellInfo>> rowDatas = new ArrayList<List<CellInfo>>();
        if (ExcelFile.readSheetData(sheet, rowDatas, 0, 1, colCount)){
            if (rowDatas.size() > 0){
                fillColWidthData(rowDatas.get(0), colWidths);
                for (CellInfo cellInfo: rowDatas.get(0)){
                    columnDatas.add(cellInfo);
                }
                return true;
            }
        }
        return false;
    }

    private static void fillColWidthData(List<CellInfo> cellInfos, List<Integer> colWidths) {
        for (int i = 0; i < cellInfos.size(); i++){
            cellInfos.get(i).setColWidth(colWidths.get(i));
        }
    }
}

/*
那就用bai下面这个du函数来执zhi行你的daobat吧。
function  WinExecAndWait32(FileName:String;  Visibility  :  integer):  DWORD;
var
        zAppName:array[0..512]  of  char;
        zCurDir:array[0..255]  of  char;
        WorkDir:String;
        StartupInfo:TStartupInfo;
        ProcessInfo:TProcessInformation;
begin
        StrPCopy(zAppName,FileName);
        GetDir(0,WorkDir);
        StrPCopy(zCurDir,WorkDir);
        FillChar(StartupInfo,Sizeof(StartupInfo),#0);
        StartupInfo.cb  :=  Sizeof(StartupInfo);
        StartupInfo.dwFlags  :=  STARTF_USESHOWWINDOW;
        StartupInfo.wShowWindow  :=  Visibility;
        if  not  CreateProcess(
        nil,
        zAppName,  {  pointer  to  command  line  string  }
        nil,  {  pointer  to  process  security  attributes  }
        nil,  {  pointer  to  thread  security  attributes  }
        false,  {  handle  inheritance  flag  }
        CREATE_NEW_CONSOLE  or  {  creation  flags  }
        NORMAL_PRIORITY_CLASS,
        nil,  {  pointer  to  new  environment  block  }
        nil,  {  pointer  to  current  directory  name  }
        StartupInfo,  {  pointer  to  STARTUPINFO  }
        ProcessInfo  {  pointer  to  PROCESS_INF  }
        )
        then  Result  :=  $FFFFFFFF  else  begin
                WaitforSingleObject(ProcessInfo.hProcess,INFINITE);
                GetExitCodeProcess(ProcessInfo.hProcess,Result);
        end;
end;
 */