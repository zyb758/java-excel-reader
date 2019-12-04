package com.zhangyibo.formatreader.format.excel;

import com.zhangyibo.formatreader.base.FormatClass;
import com.zhangyibo.formatreader.base.FormatField;
import com.zhangyibo.formatreader.base.Reader;
import com.zhangyibo.formatreader.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Iterator;
import java.util.List;

/**
 * Excel Reader impl
 *
 * @param <T>
 */
public class ExcelReader<T> implements Reader<T> {
    private int sheetIndex;
    private int rowIndex;
    private Field[] cellFields;
    private Class<T> clazz;
    private ExcelAdapter adapter;
    private boolean opened;

    @Override
    public void open(File file, Class<T> clazz) throws IOException {
        open(file, 0, clazz);
    }

    @Override
    public void open(File file, int sheetIndex, Class<T> clazz) throws IOException {
        if (file == null || clazz == null)
            throw new NullPointerException();
        if (opened)
            return;
        // instance adapter by type ...
        if (file.getName().endsWith(".xlsx")) {
            adapter = new XSSFAdapter();
        } else if (file.getName().endsWith(".xls")) {
            adapter = new HSSFAdapter();
        } else {
            throw new IOException("file name not support : " + file.getName());
        }
        this.rowIndex = 0;
        this.clazz = clazz;
        this.sheetIndex = sheetIndex;
        openExcel(file);
        configClass();
        opened = true;
    }

    private T createResult() {
        T result = null;
        try {
            result = clazz.newInstance();
        } catch (IllegalAccessException e) {
        } catch (InstantiationException e) {
        }
        return result;
    }

    private void configClass() {
        // check titles
        List<String> titles = getTitles();
        if (titles == null || titles.size() == 0) {
            throw new IllegalStateException("check excel titles ...");
        }

        // check class annotation
        if (!clazz.isAnnotationPresent(FormatClass.class)) {
            throw new IllegalStateException("class should has FormatClass");
        }

        // config fields
        cellFields = new Field[titles.size()];
        Field[] fields = clazz.getFields();
        for (Field field : fields) {
            FormatField formatField = field.getAnnotation(FormatField.class);
            int fieldIndex = titles.indexOf(formatField.value());
            if (fieldIndex >= 0) cellFields[fieldIndex] = field;
        }
    }

    private void openExcel(File file) throws IOException {
        if (adapter != null) {
            adapter.openFile(file, sheetIndex);
        }
    }

    @Override
    public List<String> getTitles() {
        if (adapter != null) {
            return adapter.getTitles();
        }
        return null;
    }

    @Override
    public T nextObject() {
        if (!opened) return null;
        final Iterator<Cell> iterator = adapter.getCells(++rowIndex);
        if (iterator == null)
            return null;
        T result = createResult();
        while (iterator.hasNext()) {
            final Cell cell = iterator.next();
            int colIndex = cell.getColumnIndex();
            Field field = cellFields[colIndex];
            if (field == null)
                continue;
            ExcelUtil.setObjValue(result, field, cell);
        }
        return result;
    }

    @Override
    public boolean isOpened() {
        return opened;
    }

    @Override
    public int index() {
        return rowIndex;
    }

    @Override
    public void close() {
        if (adapter != null) {
            adapter.close();
            adapter = null;
        }
        opened = false;
    }
}
