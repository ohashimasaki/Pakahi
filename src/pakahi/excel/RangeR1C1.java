package pakahi.excel;

import java.lang.annotation.*;

//----------------------------------------------------------------------------------------------------
/**
 * 表領域内の相対的なセルアドレスを指示するアノテーションです。
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface RangeR1C1 {
    int row() default 1;
    int column() default 1;
}
//----------------------------------------------------------------------------------------------------