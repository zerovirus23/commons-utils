package co.zero.common.files;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Date;

/**
 * Created by htenjo on 6/22/16.
 */
public class ExcelUtils {
    /**
     * @param workbook
     */
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

    /**
     *
     * @param sheet
     * @param regionIndex
     */
    private static void removeMergedRegionFillingValues(Sheet sheet, int regionIndex){
        CellRangeAddress cellRangeAddress = sheet.getMergedRegion(regionIndex);
        sheet.removeMergedRegion(regionIndex);
        setValueToCellRange(sheet, cellRangeAddress);
    }

    /**
     *
     * @param sheet
     * @param cellRangeAddress
     */
    private static void setValueToCellRange(Sheet sheet, CellRangeAddress cellRangeAddress){
        int firstRowIndex = cellRangeAddress.getFirstRow();
        int firstColumnIndex = cellRangeAddress.getFirstColumn();
        int lastRowIndex = cellRangeAddress.getLastRow();
        int lastColumIndex = cellRangeAddress.getLastColumn();
        Object cellValue;

        CellStyle filledStyle = sheet.getWorkbook().createCellStyle();
        filledStyle.setFillForegroundColor(IndexedColors.GOLD.getIndex());
        filledStyle.setFillPattern(CellStyle.SQUARES);

        for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++) {
            for (int columnIndex = firstColumnIndex; columnIndex <= lastColumIndex; columnIndex++) {
                cellValue = getValueFromMergedRegion(sheet, cellRangeAddress);
                sheet.getRow(rowIndex).getCell(columnIndex).setCellStyle(filledStyle);
                setCellValue(sheet.getRow(rowIndex).getCell(columnIndex), cellValue);
            }
        }
    }

    /**
     *
     * @param sheet
     * @param cellRangeAddress
     * @return
     */
    private static Object getValueFromMergedRegion(Sheet sheet, CellRangeAddress cellRangeAddress){
        int firstColumnIndex = cellRangeAddress.getFirstColumn();
        int firstRowIndex = cellRangeAddress.getFirstRow();
        return getCellValue(sheet.getRow(firstRowIndex).getCell(firstColumnIndex));
    }

    /**
     *
     * @param cell
     * @return
     */
    public static Object getCellValue(Cell cell){
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                return StringUtils.trim(cell.getRichStringCellValue().getString());
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_FORMULA:
                return cell.getCellFormula();
            default:
                return StringUtils.EMPTY;
        }
    }

    /**
     *
     * @param cell
     * @return
     */
    public static String getCellStringValue(Cell cell){
        return getCellValue(cell).toString();
    }

    /**
     *
     * @param cell
     * @return
     */
    public static String getCellStringValueNoSpaces(Cell cell){
        return StringUtils.trim(getCellValue(cell).toString());
    }

    /**
     *
     * @param cell
     * @return
     */
    public static double getCellNumericValue(Cell cell){
        if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
            return cell.getNumericCellValue();
        }else{
            throw new IllegalArgumentException("Cell value is not a Number");
        }
    }

    /**
     *
     * @param cell
     * @return
     */
    public static Date getCellDateValue(Cell cell){
        if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(cell)){
            return cell.getDateCellValue();
        }else{
            throw new IllegalArgumentException("Cell value is not a Date");
        }
    }

    /**
     *
     * @param cell
     * @return
     */
    public static boolean getCellBooleanValue(Cell cell){
        if(cell.getCellType() == Cell.CELL_TYPE_BOOLEAN){
            return cell.getBooleanCellValue();
        }else{
            throw new IllegalArgumentException("Cell value is not a Boolean");
        }
    }

    /**
     *
     * @param cell
     * @param value
     */
    private static void setCellValue(Cell cell, Object value){
        if (value instanceof Number){
            cell.setCellValue((Double)value);
        }else if (value instanceof Date){
            cell.setCellValue((Date)value);
        }else if (value instanceof Boolean){
            cell.setCellValue((Boolean)value);
        }else{
            cell.setCellValue((String)value);
        }
    }

    /**
     *
     * @param cell
     * @param color
     * @param pattern
     * @return
     */
    public static CellStyle buildBasicCellStyle(Cell cell, short color, short pattern){
        return buildBasicStyle(cell.getRow().getSheet().getWorkbook(), color, pattern);
    }

    /**
     *
     * @param row
     * @param color
     * @param pattern
     * @return
     */
    public static CellStyle buildBasicCellStyle(Row row, short color, short pattern){
        return buildBasicStyle(row.getSheet().getWorkbook(), color, pattern);
    }

    /**
     *
     * @param sheet
     * @param color
     * @param pattern
     * @return
     */
    public static CellStyle buildBasicCellStyle(Sheet sheet, short color, short pattern){
        return buildBasicStyle(sheet.getWorkbook(), color, pattern);
    }

    /**
     * O
     * @param workbook
     * @param color
     * @param pattern
     * @return
     */
    public static CellStyle buildBasicStyle(Workbook workbook, short color, short pattern){
        CellStyle filledStyle = workbook.createCellStyle();
        filledStyle.setFillForegroundColor(color);
        filledStyle.setFillPattern(pattern);
        return filledStyle;
    }

    public static String rowToString(Row row){
        StringBuilder rowContent = new StringBuilder();
        Cell currentCell;

        if(row != null){
            for(int i=0; i<row.getLastCellNum(); i++){
                currentCell = row.getCell(i);

                if(currentCell != null){
                    rowContent.append("| Cell " + i + " = " + getCellValue(currentCell  ).toString());
                }else{
                    rowContent.append("| Cell " + i + " = NULL");
                }
            }
        }


        return rowContent.toString();
    }
}
