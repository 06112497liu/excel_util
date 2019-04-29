package com.tellyes.platform.toolkit.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * <pre>
 * 用于标识Excel导出方法
 * 通过这个注解找到上层方法所在类型、方法名找到对应的excel导出的json配置文件名称
 * 导出的excel配置文件必须与方法所在类的同层目录且名字为className.methodName.json
 * 建议加载Controller方法上
 * </pre>
 * @author xiehai
 * @date 2018/06/12 17:44
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExportConfig {
    /**
     * 导出配置类型 默认{@link ExportConfigType#YAML}
     * @return {@link ExportConfigType}
     */
    ExportConfigType value() default ExportConfigType.YAML;
}