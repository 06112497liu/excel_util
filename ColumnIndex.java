package com.tellyes.platform.toolkit.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel导入用于标记列与字段对应关系
 * @author xiehai
 * @date 2018/09/20 15:46
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
@Target({ElementType.METHOD, ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ColumnIndex {
    /**
     * excel列索引默认从0开始
     * @return 列索引
     */
    int value();
}
