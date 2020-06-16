package com.mindmotion.xls2latex.file;

import com.mindmotion.xls2latex.common.Rect;
import com.mindmotion.xls2latex.domain.CellInfo;
import com.mindmotion.xls2latex.domain.ColorInfo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class ExcelFile {
    public static boolean readSheetData(Sheet sheet, List<List<CellInfo>> datas, int startRowIndex, int length, int colCount) {
        int rowCount = getRowCount(length, startRowIndex, sheet.getLastRowNum());
        for (int i = startRowIndex; i < rowCount; i ++) {
            if (isEmptyRow(sheet.getRow(i), colCount)) {
                break;
            }

            List<CellInfo> rowData = new ArrayList<CellInfo>();
            if (readRowData(sheet, i, colCount, rowData)){
                datas.add(rowData);
            };
        }
        return true;
    }

    private static Boolean readRowData(Sheet sheet, Integer rowIndex, int colCount, List<CellInfo> rowData) {
        Row row = sheet.getRow(rowIndex);
        Rect rect = new Rect();
        for (int i = 0; i < colCount; i ++){
            Cell cell = row.getCell(i);
            CellInfo cellInfo = new CellInfo();
            cellInfo.setMerged(isMerged(sheet, rowIndex, i, rect));
            cellInfo.setRect(rect);
            cellInfo.sethAligment(getHAligment(cell));
            cellInfo.setvAligment(getVAligment(cell));
            cellInfo.setBackColor(getBackColorFromCell(cell));
            cellInfo.setFontColor(getFontColorFromCell(cell));
            cellInfo.setText(getTextFromCell(cell).trim());
            rowData.add(cellInfo);
        }
        return true;
    }

    private static Integer getFontColorFromCell(Cell cell) {
        if (cell != null) {
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cell.getCellStyle();
            XSSFFont xssfFont = xssfCellStyle.getFont();
            return color2RGB(xssfFont.getXSSFColor());
        } else {
            return color2RGB(null);
        }
    }

    private static Integer getBackColorFromCell(Cell cell) {
        if (cell != null){
            CellStyle cellStyle = cell.getCellStyle();
            return color2RGB(cellStyle.getFillForegroundColorColor());
        } else {
            return color2RGB(null);
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

    private static VerticalAlignment getVAligment(Cell cell) {
        if (cell != null) {
            CellStyle cellStyle = cell.getCellStyle();
            return cellStyle.getVerticalAlignment();
        } else {
            return VerticalAlignment.CENTER;
        }
    }

    private static HorizontalAlignment getHAligment(Cell cell) {
        if (cell != null) {
            CellStyle cellStyle = cell.getCellStyle();
            if (cellStyle.getAlignment() == HorizontalAlignment.GENERAL) {
                return HorizontalAlignment.LEFT;
            } else {
                return cellStyle.getAlignment();
            }
        } else {
            return HorizontalAlignment.LEFT;
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
                    rect.setLeft(firstColumn);
                    rect.setTop(firstRow);
                    rect.setRight(lastColumn);
                    rect.setBottom(lastRow);
                    return true;
                }
            }
        }
        rect.setLeft(columnIndex);
        rect.setTop(rowIndex);
        rect.setRight(columnIndex);
        rect.setBottom(rowIndex);
        return false;
    }

    private static int getRowCount(int length, int startRowIndex, int lastRowNum) {
        if (lastRowNum == -1){
            return 0;
        } else {
            return length == -1 ? lastRowNum - startRowIndex + 1 : Math.min(length, lastRowNum + 1);
        }
    }

    private static boolean isEmptyRow(Row row, int colCount) {
        String text;
        for (int i = 0; i < colCount; i ++){
            text = getTextFromCell(row.getCell(i));
            if (text != null && "".equals(text.trim()) == false){
                return false;
            }
        }
        return true;
    }

    private static String getTextFromCell(Cell cell) {
        if (cell != null){
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
        } else {
            return "";
        }
    }
}
