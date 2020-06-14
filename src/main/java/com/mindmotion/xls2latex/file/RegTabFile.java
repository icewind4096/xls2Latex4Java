package com.mindmotion.xls2latex.file;

import com.mindmotion.xls2latex.common.Rect;
import com.mindmotion.xls2latex.domain.CellInfo;
import com.mindmotion.xls2latex.domain.ColorInfo;
import com.mindmotion.xls2latex.domain.ParamaterInfo;
import com.mindmotion.xls2latex.enums.ResultEnum;
import com.mindmotion.xls2latex.util.FileUtil;
import com.mindmotion.xls2latex.util.LatexUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class RegTabFile {
    public static Integer GeneralFile(ParamaterInfo paramaterInfo){
        List<List<CellInfo>> rowDatas = new ArrayList<List<CellInfo>>();
        List<String> latexDatas = new ArrayList<String>();
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(paramaterInfo.getSourceFileName()));
            for (int i = 0; i < workbook.getNumberOfSheets(); i ++){
                clearRowDatas(rowDatas);
                clearLatexDatas(latexDatas);
                if (readSheetData(workbook.getSheetAt(i), rowDatas, 0, -1, paramaterInfo.getColCount())){
                    if (translate2RegTab(paramaterInfo.getLanguage(), rowDatas, latexDatas)){
                        if (!generalRegFile(getRegTabFileName(paramaterInfo.getDestDirectory(), getRegTabFileNamePrefix(paramaterInfo.getSourceFileName()), workbook.getSheetAt(i).getSheetName()), latexDatas)){
                            return ResultEnum.MAKEOUTDIRFAIL.getCode();
                        }
                    } else {
                        return ResultEnum.READEXCELFILEFAIL.getCode();
                    }
                }
            }
            workbook.close();
            return ResultEnum.SUCCESS.getCode();
        } catch (IOException e) {
            e.printStackTrace();
            return ResultEnum.READEXCELFILEFAIL.getCode();
        }
    }

    private static boolean generalRegFile(String fileName, List<String> datas) {
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

    private static boolean readSheetData(Sheet sheet, List<List<CellInfo>> datas, int startRowIndex, int length, int colCount) {
        int rowCount = getRowCount(length, startRowIndex, sheet.getLastRowNum());
        for (int i = startRowIndex; i < rowCount; i ++) {
            if (isEmptyRow(sheet.getRow(i))) {
                break;
            }

            List<CellInfo> rowData = new ArrayList<CellInfo>();
            if (readRowData(sheet, i, colCount, rowData)){
                datas.add(rowData);
            };
        }
        return true;
    }

    private static boolean isEmptyRow(Row row) {
        String text = getTextFromCell(row.getCell(0));
        return (text == null) | ("".equals(text.trim()));
    }

    private static int getRowCount(int length, int startRowIndex, int lastRowNum) {
        if (lastRowNum == -1){
            return 0;
        } else {
            return length == -1 ? lastRowNum - startRowIndex + 1 : Math.min(length, lastRowNum + 1);
        }
    }

    private static Boolean readRowData(Sheet sheet, Integer rowIndex, int colCount, List<CellInfo> rowData) {
        Row row = sheet.getRow(rowIndex);
        Rect rect = new Rect();
        for (int i = 0; i < colCount; i ++){
            Cell cell = row.getCell(i);
            CellInfo cellInfo = new CellInfo();
            cellInfo.setMerged(isMerged(sheet, rowIndex, i, rect));
            if (cellInfo.getMerged()) {
                cellInfo.setRect(rect);
            }
            cellInfo.sethAligment(getHAligment(cell));
            cellInfo.setvAligment(getVAligment(cell));
            cellInfo.setBackColor(getBackColorFromCell(cell));
            cellInfo.setFontColor(getFontColorFromCell(cell));
            cellInfo.setText(getTextFromCell(cell).trim());
            rowData.add(cellInfo);
        }
        return true;
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

    private static String getTextFromCell(Cell cell) {
        switch (cell.getCellType()) {
            case NUMERIC:
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    return sdf.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(cell.getNumericCellValue())).toString();
                } else {
                    DataFormatter dataFormatter = new DataFormatter();
                    return dataFormatter.formatCellValue(cell);
                }
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue() + "";
            case FORMULA:
                return cell.getCellFormula() + "";
            case BLANK:
            case ERROR:
            default:
                return "";
        }
    }

    private static Integer color2RGB(Color color) {
        if (color != null){
            XSSFColor xssfColor = (XSSFColor) color;
            byte[] colors = xssfColor.getRGB();
            return new ColorInfo(colors[0], colors[1], colors[2]).toRGB();
        } else {
            return new ColorInfo(255, 255, 255).toRGB();
        }
    }

    private static Integer getFontColorFromCell(Cell cell) {
        XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cell.getCellStyle();
        XSSFFont xssfFont = xssfCellStyle.getFont();
        return color2RGB(xssfFont.getXSSFColor());
    }

    private static Integer getBackColorFromCell(Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();
        return color2RGB(cellStyle.getFillForegroundColorColor());
    }

   private static VerticalAlignment getVAligment(Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();
        return cellStyle.getVerticalAlignment();
    }

    private static HorizontalAlignment getHAligment(Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle.getAlignment() == HorizontalAlignment.GENERAL){
            return HorizontalAlignment.LEFT;
        } else {
            return cellStyle.getAlignment();
        }
    }

    private static Boolean isMerged(Sheet sheet, int rowIndex, int columnIndex, Rect rect) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for(int i = 0; i < sheetMergeCount; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if(rowIndex >= firstRow && rowIndex <= lastRow){
                if(columnIndex >= firstColumn && columnIndex <= lastColumn){
                    rect.setLeft(firstRow);
                    rect.setTop(firstColumn);
                    rect.setRight(lastColumn);
                    rect.setBottom(lastRow);
                    return true;
                }
            }
        }
        return false;
    }
}
