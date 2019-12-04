package com.zhangyibo.formatreader.format.excel;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

class HSSFAdapter implements ExcelAdapter {
    private HSSFWorkbook workbook;
    private HSSFSheet sheet;
    private List<String> titles;

    @Override
    public void openFile(File file, int sheetIndex) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        workbook = new HSSFWorkbook(fis);
        sheet = workbook.getSheetAt(sheetIndex);
        titles = null;
    }

    private void configTitles() {
        titles = new ArrayList<>();
        HSSFRow row = sheet.getRow(0);
        for (Cell cell : row) {
            titles.add(cell.getStringCellValue());
        }
    }

    @Override
    public List<String> getTitles() {
        if (titles == null && sheet != null) {
            configTitles();
        }
        return titles;
    }

    @Override
    public Iterator<Cell> getCells(int rowIndex) {
        if (sheet == null)
            return null;
        HSSFRow row = sheet.getRow(rowIndex);
        if (row == null)
            return null;
        return row.cellIterator();
    }

    @Override
    public void close() {
        if (workbook != null) {
            try {
                workbook.close();
            } catch (IOException e) {
//                e.printStackTrace();
            }
            sheet = null;
            workbook = null;
        }
    }
}
