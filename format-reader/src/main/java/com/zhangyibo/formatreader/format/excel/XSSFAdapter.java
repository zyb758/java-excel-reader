package com.zhangyibo.formatreader.format.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

class XSSFAdapter implements ExcelAdapter {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private List<String> titles;

    @Override
    public void openFile(File file, int sheetIndex) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheetAt(sheetIndex);
    }

    private void configTitles() {
        titles = new ArrayList<>();
        XSSFRow row = sheet.getRow(0);
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
        XSSFRow row = sheet.getRow(rowIndex);
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
//                I have no idea what to do here  :)
            }
            sheet = null;
            workbook = null;
        }
    }
}
