package com.tellyes.platform.toolkit.utils;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory;
import com.tellyes.core.exception.UtilException;
import com.tellyes.platform.toolkit.annotation.ExportConfig;
import com.tellyes.platform.toolkit.annotation.ExportConfigType;
import com.tellyes.platform.toolkit.vo.ExportConfigRootVo;
import com.tellyes.platform.toolkit.vo.ExportConfigVo;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.annotation.AnnotationUtils;
import org.springframework.util.ReflectionUtils;

import java.lang.reflect.Method;
import java.util.Arrays;
import java.util.Optional;

import static com.tellyes.platform.toolkit.utils.ExceptionUtil.UTIL_EXCEPTION_FUNCTION;

/**
 * Excel导出配置工具类
 * @author xiehai
 * @date 2018/07/03 11:58
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
interface ExcelConfigUtil {
    Logger LOGGER = LoggerFactory.getLogger(ExcelConfigUtil.class);
    /**
     * 获得excel导出配置
     * @return {@link ExportConfigVo}
     */
    static ExportConfigVo getExcelConfig() {
        try {
            // 业务调用方法栈
            // 找到第一个方法上标记有ExportConfig注解的方法栈
            StackTraceElement stackTrace = Arrays.stream(Thread.currentThread().getStackTrace())
                .filter(trace -> {
                    try {
                        // 调用类
                        Class<?> clazz = Class.forName(trace.getClassName());
                        // 调用方法
                        Method method = Arrays.stream(ReflectionUtils.getAllDeclaredMethods(clazz))
                            .filter(m -> m.getName().equals(trace.getMethodName()))
                            .findFirst()
                            .orElse(null);
                        return
                            // 调用方法是否标记ExportConfig
                            Optional.ofNullable(method)
                                .map(m -> AnnotationUtils.findAnnotation(m, ExportConfig.class) != null)
                                .orElse(false);
                    } catch (ClassNotFoundException e) {
                        LOGGER.error(e.getMessage(), e);
                    }

                    return false;
                }).findFirst()
                .orElseThrow(() -> new UtilException("没有找到ExportConfig标记的方法!"));
            // 业务方法类
            Class<?> clazz = Class.forName(stackTrace.getClassName());
            // yaml文件相对路径
            String yamlFilePath = String.format("%s.yml", clazz.getSimpleName());
            // 获得导出excel配置
            ObjectMapper mapper = new ObjectMapper(new YAMLFactory());
            ExportConfigRootVo rootVo = mapper.readValue(
                clazz.getResourceAsStream(yamlFilePath),
                ExportConfigRootVo.class
            );

            // 获得业务方法的导出配置
            ExportConfigVo configVo = rootVo.getRoot().get(stackTrace.getMethodName());
            // 配置完整性验证
            ExportConfigVo.validate(configVo);

            return configVo;
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw UTIL_EXCEPTION_FUNCTION.apply(e);
        }
    }

    /**
     * 获得freemarker业务方法方法栈
     * @return {@link StackTraceElement}
     */
    static StackTraceElement getFtlStackTraceElement() {
        // 业务调用方法栈
        // 找到第一个方法上标记有ExportConfig注解的方法栈
        return Arrays.stream(Thread.currentThread().getStackTrace())
            .filter(trace -> {
                try {
                    // 调用类
                    Class<?> clazz = Class.forName(trace.getClassName());
                    // 调用方法
                    Method method = Arrays.stream(ReflectionUtils.getAllDeclaredMethods(clazz))
                        .filter(m -> m.getName().equals(trace.getMethodName()))
                        .findFirst()
                        .orElse(null);
                    return
                        // 调用方法是否标记ExportConfig 且为FREEMARKER的
                        Optional.ofNullable(method)
                            .map(m -> AnnotationUtils.findAnnotation(m, ExportConfig.class))
                            .filter(annotation -> annotation.value() == ExportConfigType.FREEMARKER)
                            .isPresent();
                } catch (ClassNotFoundException e) {
                    LOGGER.error(e.getMessage(), e);
                }

                return false;
            }).findFirst()
            .orElseThrow(() -> new UtilException("没有找到@ExportConfig(ExportConfigType.FREEMARKER)标记的方法!"));
    }
}
