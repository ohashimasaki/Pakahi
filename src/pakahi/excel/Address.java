package pakahi.excel;

import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

import static java.util.regex.Pattern.CASE_INSENSITIVE;



//----------------------------------------------------------------------------------------------------
/**
 * セルのアドレスを保持・操作するクラスです。
 */
public class Address {

    int row = 1;
    int column = 1;
    String columnName = "A";

    final int A = 65;
    final int Alphabet = 26;

    final int MAX_ROW = 1048576;
    final int MAX_COLUMN = 16384;


    //------------------------------------------------------------------------------------------------
    /**
     * セルのアドレスを保持・操作するクラスのコンストラクタです。
     * @param  address セルのアドレスを表す A1 形式の文字列
     */
    public Address(String address) {

        final Pattern pattern = Pattern.compile("^([A-Z]{1,2}|[A-X][A-F][A-D])(\\d+)$", CASE_INSENSITIVE);
        Matcher m = pattern.matcher(address);

        if( ! m.find()) {
            return;
        }

        columnName = m.group(1);
        column = parseColumn(columnName);
        row = Integer.parseInt(m.group(2));

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスの行番号を整数で返します。
     * @return 行番号
     */
    public int getRow() {

        return row;

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスの列番号を整数で返します。
     * @return 列番号
     */
    public int getColumn() {

        return column;

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスの行番号と列番号を設定します。
     * @param r 行番号
     * @param c 列番号
     * @return このアドレスオブジェクト
     */
    public Address set(int r, int c) {

        setRow(r);
        setColumn(c);
        return this;

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスの行番号を設定します。
     * @param r 行番号
     * @return このアドレスオブジェクト
     */
    public Address setRow(int r) {

        row = r;

        if(row < 1) {
            row = 1;
        } else if(row > MAX_ROW) {
            row = MAX_ROW;
        }

        return this;

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスの列番号を設定します。
     * @param c 列番号
     * @return このアドレスオブジェクト
     */
    public Address setColumn(int c) {

        column = c;

        if(column < 1) {
            column = 1;
        } else if(column > MAX_COLUMN) {
            column = MAX_COLUMN;
        }

        columnName = toColumnName(column);

        return this;

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスの列名を返します。
     * @return 列名
     */
    public String getColumnName() {

        return toColumnName(column);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスを A1 形式の文字列で返します。
     * @return A1 形式のアドレス
     */
    public String getA1() {

        return columnName + row;

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスを整数で返します。
     * @return 行番号・列番号のハッシュマップ
     */
    public HashMap<String, Integer> getR1C1() {

        return new HashMap<String, Integer>() {{
            put("row", row);
            put("column", column);
        }};

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスの行番号を加算します。
     * @param n 加算する行数
     * @return このアドレスオブジェクト
     */
    public Address addRow(int n) {

        return setRow(row + n);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * このアドレスの列番号を加算します。
     * @param n 加算する列数
     * @return このアドレスオブジェクト
     */
    public Address addColumn(int n) {

        return setColumn(column + n);

    }
    //------------------------------------------------------------------------------------------------
    private int parseColumn(String column) {

        return IntStream.range(0, column.length())
                      .mapToObj(i -> ((int)column.charAt(i) - A + 1) * (int)Math.pow(Alphabet, column.length() - i - 1))
                      .reduce(Integer::sum).orElse(1);

    }
    //------------------------------------------------------------------------------------------------
    private String toColumnName(int c) {

        String t = "";

        do {
            int mod = (c % Alphabet) + A - 1;
            c = c / Alphabet;
            t = (char)mod + t;
        } while(c > 26);

        if(c > 0) {
            t = (char)(c + A - 1) + t;
        }

        return t;

    }
    //------------------------------------------------------------------------------------------------

}
//----------------------------------------------------------------------------------------------------