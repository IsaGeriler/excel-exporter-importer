package org.trupt.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellHeader {
    String headerName() default "";
    boolean isRequired() default false;
    boolean calculateAverage() default false;
    boolean calculateSum() default false;
}