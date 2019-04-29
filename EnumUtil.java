package com.tellyes.core.utils;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;

import org.springframework.util.ClassUtils;

import com.tellyes.core.annotation.Key;
import com.tellyes.core.annotation.Value;
import com.tellyes.core.exception.UtilException;

/**
 * 枚举工具类
 * @author xiehai
 * @date 2018/04/12 11:44
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
public interface EnumUtil {
    /**
     * 获得枚举的key字段
     * @param clazz 枚举Class
     * @param <E>   枚举类型
     * @return {@link Field}
     */
    @SuppressWarnings("rawtypes")
    static <E extends Enum> Field getKeyField(Class<E> clazz) {
        Objects.requireNonNull(clazz);
        return
            Arrays.stream(clazz.getDeclaredFields())
                // 找到第一个加了@Key注解的字段
                .filter(field -> field.getAnnotation(Key.class) != null)
                .findFirst()
                // 若没有则找名称为key的字段
                // 否则抛出异常
                .orElseGet(() -> {
                    try {
                        return clazz.getDeclaredField(Key.DEFAULT_KEY_NAME);
                    } catch (NoSuchFieldException e) {
                        throw new UtilException(
                            String.format("%s has either @Key annotation or name is 'key' field", clazz.getName())
                        );
                    }
                });
    }

    /**
     * 获得枚举的value字段
     * @param clazz 枚举Class
     * @param <E>   枚举类型
     * @return {@link Field}
     */
    @SuppressWarnings("rawtypes")
    static <E extends Enum> Field getValueField(Class<E> clazz) {
        return
            Arrays.stream(clazz.getDeclaredFields())
                // 找到第一个加了@Value注解的字段
                .filter(field -> field.getAnnotation(Value.class) != null)
                .findFirst()
                // 若没有则找名称为value的字段
                // 否则抛出异常
                .orElseGet(() -> {
                    try {
                        return clazz.getDeclaredField(Value.DEFAULT_VALUE_NAME);
                    } catch (NoSuchFieldException e) {
                        throw new UtilException(
                            String.format("%s has either @Value annotation or name is 'value' field", clazz.getName())
                        );
                    }
                });
    }

    /**
     * 获得枚举给定key的实例
     * @param clazz 枚举class
     * @param key   枚举key
     * @param <E>   枚举类型
     * @param <T>   key类型
     * @return 枚举实例
     */
    @SuppressWarnings("rawtypes")
	static <E extends Enum, T> E get(Class<E> clazz, T key) {
        // 获得key字段
        Field keyField = getKeyField(clazz);

        // 校验key类型和枚举声明的key类型是否一致或超类
        if (!ClassUtils.isAssignable(keyField.getType(), key.getClass())) {
            throw new UtilException(
                String.format(
                    "%s's field '%s' required %s, but %s",
                    clazz.getName(),
                    keyField.getName(),
                    keyField.getType().getName(),
                    key.getClass().getName()
                )
            );
        }

        return
            Optional.of(keyField)
                .map(field -> {
                    // 字段是否是可访问的
                    boolean flag = field.isAccessible();
                    // 设置为可访问
                    if (!flag) {
                        field.setAccessible(true);
                    }

                    E e = Arrays.stream(clazz.getEnumConstants())
                        .filter(item -> {
                            try {
                                return String.valueOf(field.get(item)).equals(String.valueOf(key));
                            } catch (IllegalAccessException ex) {
                                return false;
                            }
                        }).findFirst()
                        .orElse(null);

                    // 还原访问限制
                    if (!flag) {
                        field.setAccessible(false);
                    }

                    return e;
                })
                .orElseThrow(() ->
                    new UtilException(
                        String.format("%s not found %s=%s enum", clazz.getName(), keyField.getName(), key)
                    )
                );
    }

    /**
     * 根据枚举类型的名称获得枚举{@link EnumEntity}列表
     * @param className 枚举类型全名
     * @return {@link List}
     */
    static List<EnumEntity> toEntities(String className) {
        try {
            return toEntities(Class.forName(className).asSubclass(Enum.class));
        } catch (ClassNotFoundException e) {
            throw new UtilException(String.format("class %s not found!", className));
        } catch (ClassCastException e) {
            throw new UtilException(String.format("class %s can not cast to %s!", className, Enum.class.getName()));
        }
    }

    /**
     * 将枚举常量转为{@link EnumEntity}
     * @param clazz 枚举Class
     * @param <E>   枚举类型
     * @return {@link List}
     */
    @SuppressWarnings("rawtypes")
	static <E extends Enum> List<EnumEntity> toEntities(Class<E> clazz) {
        Field keyField = getKeyField(clazz);
        boolean isKeyAccessible = keyField.isAccessible();
        Field valueField = getValueField(clazz);
        boolean isValueAccessible = valueField.isAccessible();

        if (!isKeyAccessible) {
            keyField.setAccessible(true);
        }
        if (!isValueAccessible) {
            valueField.setAccessible(true);
        }

        List<EnumEntity> collect = Arrays.stream(clazz.getEnumConstants())
            .map(item -> {
                try {
                    return new EnumEntity(keyField.get(item), valueField.get(item));
                } catch (IllegalAccessException ignored) {
                    return null;
                }
            })
            .filter(Objects::nonNull)
            .collect(Collectors.toList());

        if (!isKeyAccessible) {
            keyField.setAccessible(false);
        }
        if (!isValueAccessible) {
            valueField.setAccessible(false);
        }

        return collect;
    }

    /**
     * 枚举字典展示实体
     */
    @SuppressWarnings("serial")
	class EnumEntity implements Serializable {
        private Object key;
        private Object value;

        EnumEntity(Object key, Object value) {
            this.key = key;
            this.value = value;
        }

        public Object getKey() {
            return key;
        }

        public void setKey(Object key) {
            this.key = key;
        }

        public Object getValue() {
            return value;
        }

        public void setValue(Object value) {
            this.value = value;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) {
                return true;
            }
            if (o == null || this.getClass() != o.getClass()) {
                return false;
            }
            EnumEntity entity = (EnumEntity) o;
            return Objects.equals(this.key, entity.key) &&
                Objects.equals(this.value, entity.value);
        }

        @Override
        public int hashCode() {
            return Objects.hash(this.key, this.value);
        }

        @Override
        public String toString() {
            return "EnumEntity{" +
                "key=" + this.key +
                ", value=" + this.value +
                '}';
        }
    }
}

