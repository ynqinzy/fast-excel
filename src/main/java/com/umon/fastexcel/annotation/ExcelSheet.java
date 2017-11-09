package com.umon.fastexcel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.hssf.util.HSSFColor;

/**
 * 
 * @className:MapperSheet
 * @description:
 * <p>
 * 工作簿信息
 * </p>
 * @author qinzy
 * @datetime:2017年11月6日
 *
 */
@Target({ ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelSheet {

	String value() default "";

	/**
	 * 工作簿名称
	 *
	 * @return
	 */
	String name() default "";

	/**
	 * 表头/首行的颜色
	 *
	 * @return
	 */
	HSSFColor.HSSFColorPredefined headColor() default HSSFColor.HSSFColorPredefined.LIGHT_GREEN;

}