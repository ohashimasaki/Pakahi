package pakahi.excel;

import java.lang.annotation.*;

//----------------------------------------------------------------------------------------------------
/**
 * セルアドレスを A1 形式で指示するアノテーションです。
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Range {
    String sheet();
    String range();
}
//----------------------------------------------------------------------------------------------------