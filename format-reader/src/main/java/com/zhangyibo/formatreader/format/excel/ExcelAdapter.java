package com.zhangyibo.formatreader.format.excel;

import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

/**
 * Excel adapter
 */
interface ExcelAdapter {
    /**
     * open excel file
     *
     * @param file       - excel file
     * @param sheetIndex - excel sheet index
     * @throws IOException - exception
     */
    void openFile(File file, int sheetIndex) throws IOException;

    /**
     * get excel titles, row 0
     *
     * @return titles
     */
    List<String> getTitles();

    /**
     * get row cells
     *
     * @param rowIndex - row index
     * @return row cells
     */
    Iterator<Cell> getCells(int rowIndex);

    /**
     * close excel file
     */
    void close();
}
