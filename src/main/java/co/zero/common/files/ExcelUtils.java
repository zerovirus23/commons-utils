package co.zero.common.files;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Created by htenjo on 6/22/16.
 */
public class ExcelUtils {

    public static void removeAllMergedCellsFromWorkbook(Workbook workbook){
        int numberOfSheets = workbook.getNumberOfSheets();
        Sheet currentSheet;

        for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
            currentSheet = workbook.getSheetAt(sheetIndex);

            while(currentSheet.getNumMergedRegions() > 0){
                removeMergedRegionFillingValues(currentSheet, 0);
            }
        }
    }

    private static void removeMergedRegionFillingValues(Sheet sheet, int regionIndex){
        CellRangeAddress cellRangeAddress = sheet.getMergedRegion(regionIndex);
        String mergedCellsValue = getValueFromMergedRegion(sheet, regionIndex);
        sheet.removeMergedRegion(regionIndex);
        setValueToCellRange(sheet, cellRangeAddress, mergedCellsValue);
    }

    private static String getValueFromMergedRegion(Sheet sheet, int regionIndex){
        CellRangeAddress cellRangeAddress = sheet.getMergedRegion(regionIndex);
        int firstColumnIndex = cellRangeAddress.getFirstColumn();
        int firstRowIndex = cellRangeAddress.getFirstRow();
        return getValueFromCell(sheet.getRow(firstRowIndex).getCell(firstColumnIndex));
    }

    private static void setValueToCellRange(Sheet sheet, CellRangeAddress cellRangeAddress, String value){
        int firstRowIndex = cellRangeAddress.getFirstRow();
        int firstColumnIndex = cellRangeAddress.getFirstColumn();
        int lastRowIndex = cellRangeAddress.getLastRow();
        int lastColumIndex = cellRangeAddress.getLastColumn();

        CellStyle filledStyle = sheet.getWorkbook().createCellStyle();
        filledStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        filledStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

        for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++) {
            for (int columnIndex = firstColumnIndex; columnIndex <= lastColumIndex; columnIndex++) {
                sheet.getRow(rowIndex).getCell(columnIndex).setCellStyle(filledStyle);
                sheet.getRow(rowIndex).getCell(columnIndex).setCellValue(value);
            }
        }
    }

    private static String getValueFromCell(Cell cell){
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue().getString();
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return Double.toString(cell.getNumericCellValue());
                }
            case Cell.CELL_TYPE_BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_FORMULA:
                return cell.getCellFormula().toString();
            default:
                return StringUtils.EMPTY;
        }
    }
}
