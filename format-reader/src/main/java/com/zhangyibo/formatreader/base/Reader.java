package com.zhangyibo.formatreader.base;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * Reader interface
 *
 * @author zhangyibo
 */
public interface Reader<T> {
    /**
     * Open Excel file sheet
     *
     * @param file  - excel file
     * @param clazz - java obj class
     * @throws IOException - exception
     */
    void open(File file, Class<T> clazz) throws IOException;

    /**
     * Open Excel file sheet
     *
     * @param file       - excel file
     * @param sheetIndex - excel file sheet index
     * @param clazz      - java obj class
     * @throws IOException - exception
     */
    void open(File file, int sheetIndex, Class<T> clazz) throws IOException;

    /**
     * get excel file titles
     *
     * @return titles
     */
    List<String> getTitles();

    /**
     * get row object
     *
     * @return object
     */
    T nextObject();

    /**
     * excel file is opened
     *
     * @return file opened
     */
    boolean isOpened();

    /**
     * current read index
     *
     * @return index
     */
    int index();

    /**
     * close excel file
     */
    void close();
}
