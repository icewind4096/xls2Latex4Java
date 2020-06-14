package com.mindmotion.xls2latex.file;

import com.mindmotion.xls2latex.domain.CellInfo;
import com.mindmotion.xls2latex.domain.ParamaterInfo;
import com.mindmotion.xls2latex.enums.GeneralFileTypeEnum;
import com.mindmotion.xls2latex.enums.HAligmentEnum;
import com.mindmotion.xls2latex.enums.ResultEnum;
import com.mindmotion.xls2latex.util.FileUtil;
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
//        targetDatas.add(getBodyData(colCount, AData, AGrdType));
        targetDatas.add("\\hlineNpx{1}");
        //表格列头对齐方式结束
        targetDatas.add("\\end{tabular}");
        //表格对齐位置结束
        targetDatas.add(getTableAligmentEnd(aligment));
        //表格结束
        targetDatas.add("\\end{table}");
        return true;
    }

    private static String getColumnData(List<CellInfo> columnDatas, GeneralFileTypeEnum generalFileType) {
        return getRowData(columnDatas, generalFileType);
    }

    private static String getRowData(List<CellInfo> columnDatas, GeneralFileTypeEnum generalFileType) {
//        for (CellInfo cellInfo: )
        return "";
    }
/*
var
    i: Integer;
begin
    Result:= '';
    for i:= 0 to AColCount - 1 do
    begin
        if ARowData^.data[i].merged = True then
        begin
            if ARowData^.data[i].rect.Bottom <> ARowData^.data[i].rect.Top then
            begin
                if (ARow = ARowData^.data[i].rect.Top) then
                    Result:= Result + Format('\multirow{%d}{%dpt}{%s}', [ARowData^.data[i].rect.Bottom - ARowData^.data[i].rect.Top + 1, ARowData^.data[i].colWidth, getCellData(ARowData^.data[i].hAligment, ARow, ARowData^.data[i].colWidth, String(ARowData^.data[i].text), ARowData^.data[i].backColor, ARowData^.data[i].fontColor, AGrdType, isAutoRef(AGrdType, i, AColCount))])
            end
            else begin
                if ARowData^.data[i].rect.Right <> ARowData^.data[i].rect.Left then
                begin
                    if (i + 1 = ARowData^.data[i].rect.Left) then
//                      Result:= Result + Format('\multicolumn{%d}{%s}{%s} ', [ARowData^.data[i].rect.Right - ARowData^.data[i].rect.Left + 1, getVLine(ARowData^.data[i].hAligment, getLeftLineVis(i + 1, ARowData), getRightLineVis(i + 1, ARowData, AColCount)), getCellData(ARow, ARowData^.data[i].colWidth, String(ARowData^.data[i].text))])
                        Result:= Result + Format('\multicolumn{%d}{%s}{%s} ', [ARowData^.data[i].rect.Right - ARowData^.data[i].rect.Left + 1, getVLine(ARowData^.data[i].hAligment, getLeftLineVis(i + 1, ARowData), getRightLineVis(i + 1, ARowData, AColCount)), getCellData(ARowData^.data[i].hAligment, ARow, getMultColumnWidth(i, ARowData), String(ARowData^.data[i].text), ARowData^.data[i].backColor, ARowData^.data[i].fontColor, AGrdType, isAutoRef(AGrdType, i, AColCount))])
                    else
                        continue;
                end;
            end;
        end
        else
            Result:= Result + Format('%s', [getCellData(ARowData^.data[i].hAligment, ARow, ARowData^.data[i].colWidth, String(ARowData^.data[i].text), ARowData^.data[i].backColor, ARowData^.data[i].fontColor, AGrdType, isAutoRef(AGrdType, i, AColCount))]);
        Result:= Result + ' &';
    end;
    Delete(Result, Result.Length, 1);

    Result:= Result + Format('\\ %s ', [getCLine(AColCount, ARow, ARowData)]) + Char($0D) + Char($0A);
 */

    private static String getColumnStyle(List<CellInfo> columnDatas, GeneralFileTypeEnum generalFileType) {
        String text = "";
        for (CellInfo cellInfo : columnDatas) {
            text = text + String.format("%sp{%dpt}<{\\%s}", getVLineVisiable(generalFileType), cellInfo.getColWidth(), getColumnAlignment(cellInfo.gethAligment()));
        }

        if (generalFileType == GeneralFileTypeEnum.GENERALFILE){
            text = text.substring(1);
        }

        return text;
    }

    private static String getColumnAlignment(HorizontalAlignment aligment) {
        //左右是不是写反了??
        switch (aligment){
            case RIGHT: return "raggedleft";
            case CENTER: return "centering";
            default: return "raggedright";
        }
    }

    private static String getVLineVisiable(GeneralFileTypeEnum generalFileType) {
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