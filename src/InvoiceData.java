import pakahi.excel.IReport;
import pakahi.excel.Range;
import pakahi.excel.RangeR1C1;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.List;



//----------------------------------------------------------------------------------------------------
/**
 * テスト用のデータオブジェクトです。
 */
public class InvoiceData implements IReport {

    // 表領域に行を挿入するときは、内部クラスをリストにします。
    public InvoiceData() {
        this.item0 = new ArrayList<>() {{
            add(new InvoiceItem("X", 1, true));
            add(new InvoiceItem("Y", 2, false));
            add(new InvoiceItem("Z", 3, true));
        }};
    }

    // セルに記入するために、シート名とセルアドレスをアノテーションとして付加します。
    // 表領域に行を挿入するときは、記入開始点となる左上のセルアドレスとします。
    @Range(sheet="Sheet1", range="C10")
    public List<InvoiceItem> item0;

    @Range(sheet="Sheet1", range="A1")
    public String item1 = "A1000";

    @Range(sheet="Sheet1", range="A3")
    public int item2 = 20;

    @Range(sheet="Sheet1", range="B3")
    public double item3 = 30.03;

    @Range(sheet="Sheet1", range="C3")
    public float item4 = 0.44f;

    @Range(sheet="Sheet1", range="D3")
    public boolean item5 = true;

    @Range(sheet="Sheet1", range="E3")
    public long item6 = 60;

    @Range(sheet="Sheet1", range="A4")
    public byte item7 = 70;

    @Range(sheet="Sheet1", range="B4")
    public short item8 = 80;

    @Range(sheet="Sheet1", range="C4")
    public LocalDate item9 = LocalDate.now(ZoneId.of("UTC+09:00"));

    @Range(sheet="Sheet1", range="D4")
    public LocalTime item10 = LocalTime.now(ZoneId.of("UTC+09:00")).withNano(0);

    @Range(sheet="Sheet1", range="E4")
    public LocalDateTime item11 = LocalDateTime.now(ZoneId.of("UTC+09:00")).withNano(0);


    //------------------------------------------------------------------------------------------------
    public class InvoiceItem implements IReport {

        public InvoiceItem(String a, int b, boolean c) {
            q1 = a;
            q2 = b;
            q3 = c;
        }

        // 表領域に記入されるデータは、相対的な列番号をアノテーションとして付加します。
        @RangeR1C1(column=1)
        public String q1;

        @RangeR1C1(column=2)
        public int q2;

        @RangeR1C1(column=3)
        public boolean q3;

    }
    //------------------------------------------------------------------------------------------------

}
//----------------------------------------------------------------------------------------------------
