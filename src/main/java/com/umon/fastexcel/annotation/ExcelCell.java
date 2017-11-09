package com.umon.fastexcel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 用于映射实体类和Excel某列名称
 *
 * @author qinzy
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelCell {

	String value() default "";

	/**
	 * 在excel文件中某列数据的名称
	 *
	 * @return 名称
	 */
	String name() default "";

	/**
	 * 在excel中列的顺序，从小到大排
	 *
	 * @return 顺序
	 */
	int sn() default 0;

	/**
	 * 时间格式化，日期类型时生效
	 *
	 * @return
	 */
	String dateformat() default "yyyy-MM-dd HH:mm:ss";

	/***
	 * 数据有效性的添加
	 * 
	 * @return
	 */
	String[] data_validity() default {};
}
