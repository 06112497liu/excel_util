package com.tellyes.platform.toolkit.utils;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.tellyes.common.utils.DateUtil;
import com.tellyes.common.utils.ReflectUtil;
import com.tellyes.common.utils.StringUtil;
import com.tellyes.core.constants.Constants;
import com.tellyes.core.exception.UtilException;
import com.tellyes.platform.toolkit.annotation.ColumnIndex;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.ReflectionUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicReference;

import static com.tellyes.common.utils.ConstantDate.SDM_YYYY_MM_DD_HH_MM_SS;
import static com.tellyes.common.utils.FieldUtil.GET;

/**
 * Excel内部工具类
 * @author xiehai
 * @date 2018/07/03 10:22
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
interface ExcelUtil {
    /**
     * 单元格最大适合宽度 超过此长度的内容将自动换行
     */
    int MAX_COLUMN_WIDTH = 3200 * 4;
    Logger LOGGER = LoggerFactory.getLogger(ExcelUtil.class);
    /**
     * 表头样式
     * @param wb {@link Workbook}
     * @return {@link CellStyle}
     */
    static CellStyle headerCellStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        // 基本样式
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        // 自动换行
        style.setWrapText(true);

        // 设置字体
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);

        return style;
    }

    /**
     * 单元格数据格式
     * @param wb {@link Workbook}
     * @return {@link CellStyle}
     */
    static CellStyle columnTypeStyle(Workbook wb) {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(wb.createDataFormat().getFormat("@"));

        return cellStyle;
    }

    /**
     * 内容样式
     * @param wb {@link Workbook}
     * @return {@link CellStyle}
     */
    static CellStyle contentCellStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        // 自动换行
        style.setWrapText(true);

        // 设置字体
        Font font = wb.createFont();
        font.setBold(false);
        style.setFont(font);

        return style;
    }

    /**
     * 自动计算最大列宽缓存作为最终最适合的列宽
     * @param cache 缓存map
     * @param cell  列信息
     */
    static void autoFitColumnWidth(Map<Integer, Integer> cache, Cell cell) {
        // 当前单元格最大宽度
        int stringWidth = cell.getStringCellValue().getBytes().length * 256 + 256 * 4;
        // 超过3200 * 4宽度自动换行
        int width = stringWidth > MAX_COLUMN_WIDTH ? MAX_COLUMN_WIDTH : stringWidth;

        // 缓存最适合的列宽度
        Optional.ofNullable(cache.get(cell.getColumnIndex()))
            .filter(i -> i > width)
            // 如果列宽小于计算的列宽或未缓存列宽 则替换最大列宽
            .orElseGet(() -> {
                cache.put(cell.getColumnIndex(), width);

                return width;
            });
    }

    /**
     * 获得字段值
     * @param object    对象实体
     * @param fieldName 字段名称
     * @return 属性值
     */
    static String getFieldValue(Object object, String fieldName) {
        if (Objects.isNull(object) || StringUtil.isEmpty(fieldName)) {
            return Constants.EMPTY;
        }

         // 如果传入对象是map 直接获取key值
        if (object instanceof Map) {
            return ((Map) object).containsKey(fieldName) ? (String) ((Map) object).get(fieldName) : Constants.EMPTY;
        }

        return
            Optional.of(fieldName.contains(Constants.POINT))
                .filter(flag -> flag)
                // 带点的说明是对象子属性
                .map(flag -> {
                    AtomicReference<Object> objectAtomicReference = new AtomicReference<>(object);
                    AtomicReference<Field> fieldAtomicReference = new AtomicReference<>();
                    AtomicReference<Class<?>> classAtomicReference = new AtomicReference<>();
                    Arrays.stream(fieldName.split(Constants.TRANSFERRED_POINT))
                        .forEach(name -> {
                            Object previous = objectAtomicReference.get();
                            // 前置实例为null 结束当前遍历
                            if (Objects.isNull(previous)) {
                                return;
                            }
                            Class<?> clazz = previous.getClass();
                            Field field = ReflectionUtils.findField(clazz, name);
                            fieldAtomicReference.set(field);
                            classAtomicReference.set(clazz);
                            objectAtomicReference.set(GET.apply(field, previous));
                        });

                    return formatCellValue(
                        objectAtomicReference.get(),
                        fieldAtomicReference.get(),
                        classAtomicReference.get()
                    );
                })
                // 可直接访问的属性
                .orElseGet(() -> {
                    try {
                        Class<?> clazz = object.getClass();
                        Field field = ReflectionUtils.findField(clazz, fieldName);

                        return formatCellValue(GET.apply(field, object), field, clazz);
                    } catch (Exception e) {
                        LOGGER.error(e.getMessage(), e);
                    }

                    return Constants.EMPTY;
                });
    }

    /**
     * 格式化excel表格数据 时间按照JsonFormat注解格式化
     * @param o     对象
     * @param field 字段
     * @param clazz 所属clazz
     * @return 字符串
     */
    static String formatCellValue(Object o, Field field, Class<?> clazz) {
        if (Objects.isNull(o)) {
            return Constants.EMPTY;
        }

        return
            Optional.of(o)
                .filter(object -> object instanceof Date)
                .map(object -> {
                    JsonFormat annotation =
                        // 先判断属性上是否有@JsonFormat注解
                        Optional.ofNullable(field.getAnnotation(JsonFormat.class))
                            // 再判断getter上是否有@JsonFormat注解
                            .orElse(ReflectUtil.findGetter4Property(clazz, field.getName())
                                .getAnnotation(JsonFormat.class));


                    return
                        Optional.ofNullable(annotation)
                            // 按照给定的pattern处理时间
                            .map(a -> DateUtil.dateToString((Date) o, annotation.pattern()))
                            // 没有JsonFormat按照yyyy-MM-dd HH:mm:ss处理
                            .orElse(DateUtil.dateToString((Date) o, SDM_YYYY_MM_DD_HH_MM_SS));
                })
                .orElse(String.valueOf(o));
    }

    /**
     * 根据excel单元格的值填充属性
     * @param o         要填充的对象
     * @param field     字段
     * @param cellValue 单元格值
     */
    static void setFieldValue(Object o, Field field, String cellValue) {
        Class<?> type = field.getType();
        Object value;
        if (type == boolean.class || type == Boolean.class) {
            value = Boolean.valueOf(cellValue);
        } else if (type == int.class || type == Integer.class) {
            value = Integer.valueOf(cellValue);
        } else if (type == short.class || type == Short.class) {
            value = Short.valueOf(cellValue);
        } else if (type == byte.class || type == Byte.class) {
            value = Byte.valueOf(cellValue);
        } else if (type == float.class || type == Float.class) {
            value = Float.valueOf(cellValue);
        } else if (type == double.class || type == Double.class) {
            value = Double.valueOf(cellValue);
        } else if (type == long.class || type == Long.class) {
            value = Long.valueOf(cellValue);
        } else if (type == BigInteger.class) {
            value = new BigInteger(cellValue);
        } else if (type == BigDecimal.class) {
            value = new BigDecimal(cellValue);
        } else if (type == Date.class) {
            // 日期格式统一为 yyyy-MM-dd HH:mm:ss
            // 根据冒号是否存在区分是date 还是datetime
            if (cellValue.contains(com.tellyes.platform.toolkit.constants.Constants.COLON)) {
                value = DateUtil.stringToDatetime(cellValue);
            } else {
                value = DateUtil.stringToDate(cellValue);
            }
        } else if (type == String.class) {
            value = cellValue;
        } else {
            throw new UtilException("excel导入暂时不支持类型" + type.getName());
        }

        try {
            MethodUtil.findSetter4Field(field).invoke(o, value);
        } catch (IllegalAccessException | InvocationTargetException e) {
            LOGGER.error(e.getMessage(), e);
            throw ExceptionUtil.UTIL_EXCEPTION_FUNCTION.apply(e);
        }
    }

    /**
     * 获得导出索引、字段缓存
     * @param clazz {@link Class}
     * @return {@link Map}
     */
    static Map<Integer, Field> getExportFields(Class<?> clazz) {
        Map<Integer, Field> columnIndexCache = new HashMap<>(8);
        Arrays.stream(FieldUtils.getAllFields(clazz))
            .forEach(field -> {
                // 属性上的@ColumnIndex
                ColumnIndex index = field.getAnnotation(ColumnIndex.class);
                int current = -1;
                if (Objects.nonNull(index)) {
                    current = index.value();
                    columnIndexCache.put(current, field);
                }

                // setter方法上存在@ColumnIndex 则以setter方法为准
                index = MethodUtil.findSetter4Property(clazz, field.getName()).getAnnotation(ColumnIndex.class);
                if (Objects.nonNull(index)) {
                    columnIndexCache.remove(current);
                    columnIndexCache.put(index.value(), field);
                }
            });

        return columnIndexCache;
    }
}
