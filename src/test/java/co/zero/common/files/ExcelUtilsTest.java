package co.zero.common.files;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URL;

import static org.junit.Assert.*;

/**
 * Created by htenjo on 6/22/16.
 */
public class ExcelUtilsTest {
    private static final String ROOT_PATH = "/";
    private static final String FILE_WITH_MERGED_CELLS_PATH = "ExcelWithMergedCells.xlsx";
    private static final String FILE_WITHOUT_MERGED_CELLS_PATH = "ExcelWithoutMergedCells.xlsx";

    @org.junit.Before
    public void setUp() throws Exception {

    }

    @org.junit.After
    public void tearDown() throws Exception {

    }

    @org.junit.Test
    public void removeAllMergedCellsFromWorkbook() throws Exception {
        URL fileUrl = getClass().getResource(ROOT_PATH + FILE_WITH_MERGED_CELLS_PATH);
        FileInputStream fis = new FileInputStream(fileUrl.getFile());
        Workbook workbook = WorkbookFactory.create(fis);
        assertEquals(3, workbook.getSheetAt(0).getNumMergedRegions());
        ExcelUtils.removeAllMergedCellsFromWorkbook(workbook);
        assertEquals(0, workbook.getSheetAt(0).getNumMergedRegions());
        fis.close();

        URL rootUrl = getClass().getResource(ROOT_PATH);
        File file = new File(rootUrl.getFile() + FILE_WITHOUT_MERGED_CELLS_PATH);
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        fos.close();
    }
}