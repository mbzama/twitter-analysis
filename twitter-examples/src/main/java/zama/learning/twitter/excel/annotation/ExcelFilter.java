package zama.learning.twitter.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD})
public @interface ExcelFilter {
	/**
	 * @return The title you want to display
	 */
	String title();
	/**
	 * 
	 * @return The row you want this item to be
	 */
	int row() default 0;
	/**
	 * 
	 * @return The order in the row you want this item to be.
	 */
	int order() default 0;
}
