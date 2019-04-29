package com.tellyes.platform.toolkit.utils;

import com.tellyes.common.utils.CollectionUtil;
import com.tellyes.common.utils.StringUtil;
import com.tellyes.core.config.ThreadPoolConfig;
import com.tellyes.core.constants.Constants;
import com.tellyes.core.constants.FileTypeEnum;
import com.tellyes.core.exception.UtilException;
import com.tellyes.core.plugin.sequence.UUIDSequence;
import com.tellyes.core.utils.DownloadUtil;
import com.tellyes.core.utils.SpringUtil;
import com.tellyes.platform.toolkit.vo.ExportConfigVo;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.annotation.Import;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.IntStream;

import static com.tellyes.platform.toolkit.utils.ExcelConfigUtil.getExcelConfig;
import static com.tellyes.platform.toolkit.utils.ExcelUtil.autoFitColumnWidth;
import static com.tellyes.platform.toolkit.utils.ExcelUtil.columnTypeStyle;
import static com.tellyes.platform.toolkit.utils.ExcelUtil.headerCellStyle;
import static com.tellyes.platform.toolkit.utils.ExceptionUtil.UTIL_EXCEPTION_FUNCTION;

/**
 * Excel通用导出工具类
 * @author xiehai
 * @date 2018/06/08 08:54
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
@Import(ThreadPoolConfig.class)
public class ExportExcelUtil {
    /**
     * 临时excel文件存放位置
     */
    private static final String TEMP_EXCEL_PATH = "/WEB-INF/temp";
    private static final Logger LOGGER = LoggerFactory.getLogger(ExportExcelUtil.class);
    /**
     * 线程池
     */
    private static final ThreadPoolTaskExecutor EXECUTOR = SpringUtil.getBean("threadPoolTaskExecutor");

    private ExportExcelUtil() {
    }

    /**
     * 临时保存excel文件 配置为json
     * @param list 数据
     * @return 文件名
     */
    public static String save(List<?> list) {
        return save(list, getExcelConfig());
    }

    /**
     * 保存临时excel文件
     * @param list     数据
     * @param configVo 自定义配置
     * @return 文件名
     */
    public static String save(List<?> list, ExportConfigVo configVo) {
        SXSSFWorkbook workbook = initWorkbook(configVo, list);

        // 临时文件名
        String fileName = String.format("%s_%s", configVo.getFileName(), UUIDSequence.syncUuid());

        FileOutputStream out = null;
        try {
            // 文件全路径
            File file = new File(getFileFullPath(String.format("%s.%s", fileName, FileTypeEnum.XLS.getType())));
            // temp目录不存在则创建目录
            if (!file.getParentFile().exists()) {
                file.getParentFile().mkdir();
            }

            // 文件不存在创建文件
            if (!file.exists()) {
                file.createNewFile();
            }

            out = new FileOutputStream(file);
            workbook.write(out);

            return fileName;
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw UTIL_EXCEPTION_FUNCTION.apply(e);
        } finally {
            IOUtils.closeQuietly(out);
        }
    }

    /**
     * 根据yml配置直接下载表头excel文件
     */
    public static void download() {
        download(null, getExcelConfig());
    }

    /**
     * 根据yml配置直接下载excel
     * @param list 导出数据
     */
    public static void download(List<?> list) {
        download(list, getExcelConfig());
    }

    /**
     * 根据{@code configVo}配置直接下载excel
     * @param list     数据
     * @param configVo excel配置信息
     */
    public static void download(List<?> list, ExportConfigVo configVo) {
        SXSSFWorkbook workbook = initWorkbook(configVo, list);
        // 输出文件流到response
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            workbook.write(out);
            ByteArrayInputStream inputStream = new ByteArrayInputStream(out.toByteArray());
            DownloadUtil.download(
                ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest(),
                ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse(),
                configVo.getFileName(),
                inputStream,
                FileTypeEnum.XLS.getType()
            );
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw UTIL_EXCEPTION_FUNCTION.apply(e);
        }
    }

    /**
     * 根据excel文件名下载
     * @param fileName 文件名不包括后缀
     */
    public static void download(String fileName) {
        String fileFullPath = getFileFullPath(String.format("%s.%s", fileName, FileTypeEnum.XLS.getType()));
        try {
            FileInputStream inputStream = new FileInputStream(fileFullPath);
            DownloadUtil.download(
                ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest(),
                ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse(),
                fileName.split(Constants.UNDER_SCORE)[0],
                inputStream,
                FileTypeEnum.XLS.getType()
            );
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new UtilException(e.getMessage());
        } finally {
            // 异步删除文件
            EXECUTOR.execute(() ->
                Optional.of(new File(fileFullPath))
                    // 文件是否存在
                    .filter(File::exists)
                    // 删除文件
                    .filter(File::delete)
                    .ifPresent(file -> LOGGER.debug(String.format("file %s deleted!", fileFullPath)))
            );
        }
    }

    /**
     * 获得临时文件全路径
     * @param fileName 临时文件名包含后缀
     * @return 全路径文件名
     */
    static String getFileFullPath(String fileName) {
        // 临时文件路径
        String fullPath = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest()
            .getServletContext()
            .getRealPath(TEMP_EXCEL_PATH);
        return fullPath + File.separator + fileName;
    }

    /**
     * 通过excel导出配置及数据写入excel
     * @param configVo 配置信息
     * @param data     数据内容
     * @return {@link SXSSFWorkbook}
     */
    private static SXSSFWorkbook initWorkbook(ExportConfigVo configVo, List<?> data) {
        return initWorkbook(new SXSSFWorkbook(), configVo, data);
    }

    /**
     * 通过excel导出配置及数据写入excel 用于多个sheet
     * @param workbook 指定workbook
     * @param configVo 配置信息
     * @param data     数据内容
     * @return {@link SXSSFWorkbook}
     */
    private static SXSSFWorkbook initWorkbook(SXSSFWorkbook workbook, ExportConfigVo configVo, List<?> data) {
        SXSSFSheet sheet = workbook.createSheet();
        // 存储适合的列宽
        Map<Integer, Integer> fitColumnWidth = new HashMap<>(8);

        // 表头单元格格式
        CellStyle headerStyle = headerCellStyle(workbook);
        // 绘制表头
        configVo.getHeaders()
            .forEach(headers -> {
                Row row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                // 表头
                headers.forEach(head -> {
                    Cell cell = row.createCell(row.getPhysicalNumberOfCells());
                    // 合并表头单元格
                    if (head.isMerged()) {
                        String[] indices = head.getMergeIndex().split(Constants.COMMA);
                        sheet.addMergedRegion(
                            new CellRangeAddress(
                                Integer.parseInt(StringUtil.trim(indices[0])),
                                Integer.parseInt(StringUtil.trim(indices[1])),
                                Integer.parseInt(StringUtil.trim(indices[2])),
                                Integer.parseInt(StringUtil.trim(indices[3]))
                            )
                        );
                    }
                    cell.setCellValue(head.getName());
                    cell.setCellStyle(headerStyle);
                    autoFitColumnWidth(fitColumnWidth, cell);
                });
            });

        // 获得内容数据
        Optional.ofNullable(data)
            .filter(CollectionUtil::isNotEmpty)
            .ifPresent(list -> {
                // 内容样式
                CellStyle columnStyle = ExcelUtil.contentCellStyle(workbook);
                // 写入excel内容
                list.forEach(item -> {
                    Row row = sheet.createRow(sheet.getPhysicalNumberOfRows());
                    configVo.getFields()
                        .forEach(field -> {
                            Cell cell = row.createCell(row.getPhysicalNumberOfCells());
                            cell.setCellValue(ExcelUtil.getFieldValue(item, field));
                            cell.setCellStyle(columnStyle);
                            autoFitColumnWidth(fitColumnWidth, cell);
                        });
                });
            });

        // 计算最大列数
        AtomicInteger columns = new AtomicInteger();
        configVo.getHeaders().get(0)
            .forEach(head -> {
                // 合并单元格累加
                Optional.of(head.isMerged())
                    .filter(t -> t)
                    .ifPresent(t -> {
                        String[] indices = head.getMergeIndex().split(Constants.COMMA);
                        int start = Integer.parseInt(StringUtil.trim(indices[0]));
                        int end = Integer.parseInt(StringUtil.trim(indices[1]));
                        columns.getAndAdd(end - start);
                    });
                // 非合并单元格
                columns.incrementAndGet();
            });

        // 单元格数据格式
        CellStyle cellStyle = columnTypeStyle(workbook);

        IntStream.range(0, columns.get())
            .forEach(index -> {
                // 自适应列宽
                sheet.setColumnWidth(index, fitColumnWidth.get(index));
                // 设置单元格数据格式
                sheet.setDefaultColumnStyle(index, cellStyle);
            });

        // 冻结表头
        Optional.ofNullable(configVo.getFreezeIndex())
            .filter(StringUtil::isNotEmpty)
            .filter(s -> s.contains(Constants.COMMA))
            .map(t -> {
                String[] indices = t.split(Constants.COMMA);
                // 若设置了冻结索引按照设置冻结窗口
                sheet.createFreezePane(
                    Integer.parseInt(StringUtil.trim(indices[0])),
                    Integer.parseInt(StringUtil.trim(indices[1]))
                );

                return Constants.EMPTY;
            })
            .orElseGet(() -> {
                // 默认冻结行
                sheet.createFreezePane(0, configVo.getHeaders().size());

                return Constants.EMPTY;
            });

        return workbook;
    }
}
