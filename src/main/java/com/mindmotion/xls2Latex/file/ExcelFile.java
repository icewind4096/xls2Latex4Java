package com.mindmotion.xls2Latex.file;

import com.mindmotion.xls2Latex.common.Rect;
import com.mindmotion.xls2Latex.domain.CellInfo;
import com.mindmotion.xls2Latex.domain.ColorInfo;
import com.mindmotion.xls2Latex.domain.ParamaterInfo;
import com.mindmotion.xls2Latex.util.FileUtil;
import com.mindmotion.xls2Latex.util.LatexUtil;
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
import java.util.regex.Pattern;

public class ExcelFile {
    public static Boolean ProduceRegTab(ParamaterInfo paramaterInfo){
        List<List<CellInfo>> datas = new ArrayList<List<CellInfo>>();
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(paramaterInfo.getSourceFileName()));
            for (int i = 0; i < workbook.getNumberOfSheets(); i ++){
                clearDataList(datas);
                if (readSheetData(workbook.getSheetAt(i), datas, 0, -1, paramaterInfo.getColCount()) == true){
                    if (translate2RegTab(paramaterInfo.getLanguage(), paramaterInfo.getDestFileName(), datas) == false){
                    }
                }
            }
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    private static boolean translate2RegTab(int language, String fileName, List<List<CellInfo>> datas) {
        List<String> lists = new ArrayList<String>();

        generalRegHead(lists, language);

        generalRegBody(lists, datas);

        return FileUtil.saveToFileByList(fileName, lists);
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
            if (isEmptyRow(sheet.getRow(i)) == false){
                List<CellInfo> rowData = new ArrayList<CellInfo>();
                if (readRowData(sheet, i, colCount, rowData) == true){
                    datas.add(rowData);
                };
            } else {
                break;
            }
        }
        return true;
    }

    private static boolean isEmptyRow(Row row) {
        String text = getTextFromCell(row.getCell(0));
        return (text == null) | (text.trim() == "");
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
            if (cellInfo.getMerged() == true) {
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

    private static void clearDataList(List<List<CellInfo>> datas) {
        for (int i = 0; i < datas.size() - 1; i++ ){
            clearRowData(datas.get(i));
        }
        datas.clear();
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

    private static Integer getFontColorFromCell(Cell cell) {
        XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cell.getCellStyle();
        XSSFFont xssfFont = xssfCellStyle.getFont();
        return color2UOF(xssfFont.getXSSFColor());
    }

    private static Integer color2UOF(Color color) {
        if (color != null){
            XSSFColor xssfColor = (XSSFColor) color;
            byte[] colors = xssfColor.getRGB();
            return new ColorInfo(colors[0], colors[1], colors[2]).toRGB();
        } else {
            return new ColorInfo(255, 255, 255).toRGB();
        }
    }

    private static Integer getBackColorFromCell(Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();
        return color2UOF(cellStyle.getFillForegroundColorColor());
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

/*


    private static Integer getBackColorFromCell(Cell cell) {
    }

    private static String getTextFromCell(Cell cell) {
    }

    private static Boolean isMerged(Sheet sheet, int rowIndex, int columnIndex) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for(int i = 0; i < sheetMergeCount; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if(rowIndex == firstRow && rowIndex == lastRow){
                if(columnIndex >= firstColumn && columnIndex <= lastColumn){
                    return false;
                }
            }
        }
        return null ;
    }

    private static void generatorRegTabBoday(Sheet sheet, Integer colCount, StringBuilder stringBuilder) {
        generatorRegTabBodyHeadRang(stringBuilder);
        for (int i = 0; i < sheet.getLastRowNum(); i ++){

        }
        generatorRegTabBodyEndRang(stringBuilder);
    }

    private static void generatorRegTabBodyEndRang(StringBuilder stringBuilder) {
        stringBuilder.append("}");
        appendEnterLine(stringBuilder);
    }

    private static void generatorRegTabBodyHeadRang(StringBuilder stringBuilder) {
        stringBuilder.append('{');
        appendEnterLine(stringBuilder);
    }

    private static void generatorRegTabHead(Integer language, StringBuilder stringBuilder) {
        if (language == 0) {
            stringBuilder.append("regDescriptionCN");
        } else {
            stringBuilder.append("regDescriptionEN");
        }
        appendEnterLine(stringBuilder);
    }

    private static void appendEnterLine(StringBuilder stringBuilder) {
        stringBuilder.append("/r/n");
    }

    private static void translate2RegTab(Sheet sheet, Integer language, Integer colCount) {
        StringBuilder stringBuilder = new StringBuilder();
        generatorRegTabHead(language, stringBuilder);
        generatorRegTabBoday(sheet, colCount, stringBuilder);
//        for (int i = 0; i < sheet.getLastRowNum(); i++){

//        }
//        List<ColumnStyleInfo> columnsStyle = getColumnsStyleFromSheet(sheet, colWidths.size());
    }

 */
