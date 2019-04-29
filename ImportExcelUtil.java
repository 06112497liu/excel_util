package com.tellyes.platform.toolkit.utils;

import com.tellyes.common.utils.DateUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Excel导入工具类
 * @author xiehai
 * @date 2018/09/20 14:09
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
public interface ImportExcelUtil {
    /**
     * 解析excel文件第一个sheet 从头第一行数据开始解析到实体
     * @param file  excel文件
     * @param clazz 实体Class
     * @param <T>   实体类型
     * @return {@link LinkageError}
     */
    static <T> List<T> parse(File file, Class<T> clazz) {
        // 从第二行开始解析
        return parse(1, file, clazz);
    }

    /**
     * 解析excel文件第一个sheet 从给定行开始解析到实体
     * @param startRow 解析开始行数
     * @param file     excel文件
     * @param clazz    实体Class
     * @param <T>      实体类型
     * @return {@link List}
     */
    static <T> List<T> parse(int startRow, File file, Class<T> clazz) {
        // 默认解析第一个sheet
        return parse(0, startRow, file, clazz);
    }

    /**
     * 解析excel文件给定的sheet 从给定行开始解析到实体
     * @param sheetIndex sheet索引
     * @param startRow   解析开始行数
     * @param file       excel文件
     * @param clazz      实体Class
     * @param <T>        实体类型
     * @return {@link List}
     */
    static <T> List<T> parse(int sheetIndex, int startRow, File file, Class<T> clazz) {
        Objects.requireNonNull(file);
        Objects.requireNonNull(clazz);

        List<T> list = new ArrayList<>();
        if (sheetIndex < 0 || startRow < 0) {
            return list;
        }

        try {
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            if (sheet.getPhysicalNumberOfRows() < startRow) {
                return list;
            }

            Optional.of(ExcelUtil.getExportFields(clazz))
                .filter(map -> map.size() > 0)
                .ifPresent(map ->
                    sheet.forEach(row -> {
                        if (row.getRowNum() < startRow) {
                            return;
                        }
                        try {
                            T o = clazz.newInstance();
                            row.forEach(col ->
                                Optional.ofNullable(map.get(col.getColumnIndex()))
                                    // 过滤单元格为空的情况
                                    .filter(field -> col.getCellType() != Cell.CELL_TYPE_BLANK)
                                    .ifPresent(field -> {
                                        // 将时间类型转为字符串
                                        if (col.getCellType() == Cell.CELL_TYPE_NUMERIC &&
                                            HSSFDateUtil.isCellDateFormatted(col)) {
                                            col.setCellValue(DateUtil.datetimeToString(col.getDateCellValue()));
                                        }

                                        // 单元格强制设置为字符串类型
                                        col.setCellType(Cell.CELL_TYPE_STRING);
                                        ExcelUtil.setFieldValue(o, field, col.getStringCellValue());
                                    })
                            );
                            list.add(o);
                        } catch (InstantiationException | IllegalAccessException e) {
                            throw ExceptionUtil.UTIL_EXCEPTION_FUNCTION.apply(e);
                        }
                    })
                );


            return list;
        } catch (IOException | InvalidFormatException e) {
            throw ExceptionUtil.UTIL_EXCEPTION_FUNCTION.apply(e);
        }
    }
}
