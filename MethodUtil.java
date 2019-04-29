package com.tellyes.platform.toolkit.utils;

import com.tellyes.core.exception.UtilException;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * 反射方法工具类
 * @author xiehai
 * @date 2018/09/21 09:34
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
public interface MethodUtil {
    /**
     * 获得指定属性的getter方法
     * @param clazz        class
     * @param propertyName 属性名
     * @return {@link Method}
     */
    static Method findGetter4Property(Class<?> clazz, String propertyName) {
        try {
            return new PropertyDescriptor(propertyName, clazz).getReadMethod();
        } catch (IntrospectionException e) {
            throw new UtilException(e.getMessage());
        }
    }

    /**
     * 获得字段的读方法
     * @param field {@link Field}
     * @return {@link Method}
     */
    static Method findGetter4Field(Field field) {
        return findSetter4Property(field.getDeclaringClass(), field.getName());
    }


    /**
     * 获得指定属性的setter方法
     * @param clazz        class
     * @param propertyName 属性名
     * @return {@link Method}
     */
    static Method findSetter4Property(Class<?> clazz, String propertyName) {
        try {
            return new PropertyDescriptor(propertyName, clazz).getWriteMethod();
        } catch (IntrospectionException e) {
            throw new UtilException(e.getMessage());
        }
    }

    /**
     * 获得指定属性的写方法
     * @param field {@link Field}
     * @return {@link Method}
     */
    static Method findSetter4Field(Field field) {
        return findSetter4Property(field.getDeclaringClass(), field.getName());
    }
}
