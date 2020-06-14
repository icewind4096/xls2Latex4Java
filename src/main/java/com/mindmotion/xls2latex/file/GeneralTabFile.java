package com.mindmotion.xls2latex.file;

import com.mindmotion.xls2latex.domain.CellInfo;
import com.mindmotion.xls2latex.domain.ParamaterInfo;
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
                translate2GeneralTab(paramaterInfo.getTableName(), paramaterInfo.getTableName(), paramaterInfo.getColCount(), paramaterInfo.getLanguage(), paramaterInfo.getAligment(),  columnDatas, rowDatas, latexDatas);
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

    private static boolean translate2GeneralTab(String tableName, String tableLabel, Integer colCount, Integer language, Integer aligment, List<CellInfo> columnDatas, List<List<CellInfo>> rowDatas, List<String> targetDatas) {
        targetDatas.add("\\captionsetup[table]{");
        //label左对齐
        targetDatas.add("	singlelinecheck=false,");
        targetDatas.add("}");

        //设置表格label与上一段文字的距离
        targetDatas.add("\\setlength\\intextsep{10pt}");
        targetDatas.add("\\makeatletter");
        targetDatas.add("\\newcommand{\\hlineNpx}[1]{\\noalign {\\ifnum 0=`}\\fi \\hrule height #1pt	\\futurelet \\reserved@a \\@xhline}");
//        targetDatas.add(getFontSetByLanguage(ALanguage));

        //设置行间距
        targetDatas.add("\\linespread{1.2}");

        //表格开始
        targetDatas.add("\\begin{table}[h!]");
        //表格和label之间的距离
        targetDatas.add("\\setlength{\\abovecaptionskip}{1pt}");
        //表格和label之间的距离
        targetDatas.add("\\setlength{\\belowcaptionskip}{10pt}");
        //表格对齐位置开始
//        targetDatas.add(getTableAligmentBegin(AAligment));
        //表格标题
        targetDatas.add(String.format("\\caption{%s}", tableName));
        //表格引用标号
        targetDatas.add(String.format("\\label{tab:%s}", tableLabel));
        //表格列头对齐方式开始
//        targetDatas.add(String.format("\\begin{tabular}{%s}", [getColumnStyle(AData, AGrdType)]));
        targetDatas.add("\\hlineNpx{1}");
        //得到表格数据
//        targetDatas.add(getBodyData(AColCount, AData, AGrdType));
        targetDatas.add("\\hlineNpx{1}");
        //表格列头对齐方式结束
        targetDatas.add("\\end{tabular}");
        //表格对齐位置结束
//        targetDatas.add(getTableAligmentEnd(AAligment));
        //表格结束
        targetDatas.add("\\end{table}");
        return false;
    }

    private static boolean readSheetRowData(Sheet sheet, List<List<CellInfo>> rowDatas, Integer colCount, List<Integer> colWidths) {
        if (ExcelFile.readSheetData(sheet, rowDatas, 1, -1, colCount) == true){
            for (int i = 0; i < rowDatas.size(); i ++){
                fillColWidthData(rowDatas.get(i), colWidths);
            }
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