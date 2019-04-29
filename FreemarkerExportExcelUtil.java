package com.tellyes.platform.toolkit.utils;

import com.tellyes.core.constants.Constants;
import com.tellyes.core.plugin.sequence.UUIDSequence;
import com.tellyes.platform.toolkit.freemarker.ExcelDirective;
import com.tellyes.platform.toolkit.freemarker.ExcelOptionsDirective;
import com.tellyes.platform.toolkit.freemarker.ExcelStyleDirective;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateDirectiveModel;
import freemarker.template.TemplateException;
import freemarker.template.TemplateModelException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import static com.tellyes.core.constants.EncodingTypeEnum.UTF_8;
import static com.tellyes.core.constants.FileTypeEnum.FTL;
import static com.tellyes.core.constants.FileTypeEnum.XLS;
import static freemarker.template.Configuration.VERSION_2_3_23;

/**
 * freemarker导出动态列
 * @author xiehai
 * @date 2018/07/06 11:04
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
public interface FreemarkerExportExcelUtil {
    /**
     * 缓存模版
     */
    Map<String, Template> TEMPLATES = new HashMap<>(8);
    Logger LOGGER = LoggerFactory.getLogger(FreemarkerExportExcelUtil.class);

    /**
     * 保存excel
     * @param fileName 文件名
     * @param params   参数
     * @return 保存后的文件名
     */
    static String save(String fileName, Object params) {
        Map<String, Object> root = new HashMap<>(8);
        root.put("params", params);
        String tempFileName = String.format("%s_%s", fileName, UUIDSequence.syncUuid());

        try {
            File file = new File(ExportExcelUtil.getFileFullPath(tempFileName + XLS.suffix()));
            FreemarkerExportExcelUtil.getTemplate().process(
                // 传递的参数
                root,
                // 指定文件编码
                new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file), UTF_8.getType()))
            );
        } catch (TemplateException | IOException e) {
            LOGGER.error(e.getMessage(), e);
            throw ExceptionUtil.UTIL_EXCEPTION_FUNCTION.apply(e);
        }

        return tempFileName;
    }

    /**
     * 直接导出excel
     * @param fileName 导出文件名
     * @param params   参数
     */
    static void export(String fileName, Object params) {
        ExportExcelUtil.download(save(fileName, params));
    }

    /**
     * 获得对应的freemarker模版
     * @return {@link Template}
     */
    static Template getTemplate() {
        StackTraceElement ftlStackTraceElement = ExcelConfigUtil.getFtlStackTraceElement();
        try {
            Class<?> clazz = Class.forName(ftlStackTraceElement.getClassName());
            String cacheKey = clazz.getName() + Constants.POINT + ftlStackTraceElement.getMethodName() + FTL.suffix();

            return
                // 若缓存中存在 直接返回
                Optional.ofNullable(TEMPLATES.get(cacheKey))
                    // 否则将模版加入缓存
                    .orElseGet(() -> {
                        Configuration conf = new Configuration(VERSION_2_3_23);
                        conf.setClassForTemplateLoading(clazz, "");
                        conf.setDefaultEncoding(UTF_8.getType());
                        try {
                            conf.setSharedVaribles(FreemarkerExportExcelUtil.getDirectives());
                        } catch (TemplateModelException e) {
                            LOGGER.error(e.getMessage(), e);
                        }
                        Template template = null;
                        try {
                            // 模版名称
                            String templateName = clazz.getSimpleName() + Constants.POINT +
                                ftlStackTraceElement.getMethodName() + FTL.suffix();
                            // 获得指定编码的模版
                            template = conf.getTemplate(templateName, UTF_8.getType());
                            TEMPLATES.put(cacheKey, template);
                        } catch (IOException e) {
                            LOGGER.error(e.getMessage(), e);
                        }

                        return template;
                    });
        } catch (ClassNotFoundException e) {
            LOGGER.error(e.getMessage(), e);
            throw ExceptionUtil.UTIL_EXCEPTION_FUNCTION.apply(e);
        }
    }

    /**
     * 公用指令
     * @return {@link Map}
     */
    static Map<String, TemplateDirectiveModel> getDirectives() {
        // 公用指令
        Map<String, TemplateDirectiveModel> directives = new HashMap<>(8);
        // xml公共部分
        directives.put("excel", new ExcelDirective());
        // 自定义样式
        directives.put("style", new ExcelStyleDirective());
        // 冻结excel选项
        directives.put("options", new ExcelOptionsDirective());

        return directives;
    }
}
