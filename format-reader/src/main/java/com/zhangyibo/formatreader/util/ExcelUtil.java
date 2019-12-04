package com.zhangyibo.formatreader.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;

public class ExcelUtil {
    private ExcelUtil() {
    }

    public static double readCellNumericValue(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            return Double.parseDouble(cell.getStringCellValue());
        }
        throw new IllegalStateException("failed get numeric from : " + cell.getCellType());
    }

    public static String readCellStringValue(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            java.text.NumberFormat nf = java.text.NumberFormat.getInstance();
            nf.setMinimumFractionDigits(0);
            nf.setGroupingUsed(false);
            return String.valueOf(nf.format(cell.getNumericCellValue()));
        }
        throw new IllegalStateException("failed get String from : " + cell.getCellType());
    }

    private static RuntimeException getSetException(Field field, Cell cell) {
        final String cause = String.format("failed to set cell type %s to object type %s", cell.getCellType(), field.getType().getTypeName());
        return new IllegalStateException(cause);
    }

    public static void setObjValue(Object obj, Field field, Cell cell) {
        final String typeName = field.getType().getTypeName();
        if ("int".equals(typeName) || "java.lang.Integer".equals(typeName)) {
            try {
                field.setInt(obj, (int) ExcelUtil.readCellNumericValue(cell));
            } catch (IllegalAccessException e) {
                throw getSetException(field, cell);
            }
            return;
        }
        if ("double".equals(typeName) || "java.lang.Double".equals(typeName)) {
            try {
                field.setDouble(obj, ExcelUtil.readCellNumericValue(cell));
            } catch (IllegalAccessException e) {
                throw getSetException(field, cell);
            }
            return;
        }
        if ("flout".equals(typeName) || "java.lang.Flout".equals(typeName)) {
            try {
                field.setFloat(obj, (float) ExcelUtil.readCellNumericValue(cell));
            } catch (IllegalAccessException e) {
                throw getSetException(field, cell);
            }
            return;
        }
        if ("java.util.Date".equals(typeName)) {
            try {
                field.set(obj, cell.getDateCellValue());
            } catch (IllegalAccessException e) {
                throw getSetException(field, cell);
            }
            return;
        }
        if ("java.lang.String".equals(typeName)) {
            try {
                field.set(obj, ExcelUtil.readCellStringValue(cell));
            } catch (IllegalAccessException e) {
                throw getSetException(field, cell);
            }
            return;
        }
        throw new IllegalStateException("not supported type " + typeName);
    }
}
