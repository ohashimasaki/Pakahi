package pakahi.excel;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.StringWriter;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.chrono.ChronoLocalDate;
import java.time.temporal.ChronoUnit;
import java.util.Arrays;
import java.util.List;




//----------------------------------------------------------------------------------------------------
/**
 * ワークシート用の XML ファイル（/xl/worksheets/sheetN.xml）を操作します。
 */
public class Worksheet {

    File file;
    Document xml;
    SharedStrings sharedStrings;
    NamespaceContext context;
    String uri;



    //------------------------------------------------------------------------------------------------
    /**
     * ワークシート用の XML ファイル（/xl/worksheets/sheetN.xml）を操作するクラスのコンストラクタです。
     * @param tmp 一時フォルダのパス
     * @param sheetName シート名
     * @param sharedStrings 共有文字列オブジェクト
     */
    public Worksheet(Path tmp, String sheetName, SharedStrings sharedStrings) {

        this.sharedStrings = sharedStrings;

        context = new ExcelNamespaceContext();
        uri = context.getNamespaceURI("x");

        file = Paths.get(tmp.toString(), "xl/worksheets", sheetName + ".xml").toFile();

        if( ! file.exists()) {
            System.out.println("Worksheet \"" + sheetName + "\" not found: " + file.getPath());
            return;
        }

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            xml = factory.newDocumentBuilder().parse(file);

        } catch(Exception ex) {
            System.out.println(ex.getMessage());
        }

    }
    // ------------------------------------------------------------------------------------------------
    /**
     * 表領域の指定番号の行以降に行を挿入し、各セルにデータを記入します。
     * @param from 記入を開始する行番号
     * @param c 挿入する行数
     */
    public void insertRows(int from, int c) {

        Element where;
        Element source = XPath.selectSingleNode(xml, "/x:worksheet/x:sheetData/x:row[@r = " + from + "]");
        int p = 1;

        if(source == null) {
            p = 0;
            source = xml.createElementNS(uri, "row");
        }

        List<Element> rows = XPath.selectNodes(xml, "/x:worksheet/x:sheetData/x:row[@r > " + from + "]");

        if(rows != null && rows.size() > 0) {
            for(Element row : rows) {
                int r = Integer.parseInt(row.getAttribute("r")) + c - 1;
                shiftRow(row, r);
            }
            where = rows.get(0);
            for(int i = p; i < c; i++) {
                where.getParentNode().insertBefore(shiftRow((Element)source.cloneNode(true), from + i), where);
            }
        } else {
            where = XPath.selectSingleNode(xml, "/x:worksheet/x:sheetData");
            if(where != null) {
                for(int i = p; i < c; i++) {
                    where.appendChild(shiftRow((Element)source.cloneNode(true), from + i));
                }
            }
        }

    }
    //------------------------------------------------------------------------------------------------
    private Element shiftRow(Element row, int r) {

        row.setAttribute("r", String.valueOf(r));

        NodeList cells = row.getElementsByTagName("c");

        for(int j = 0; j < cells.getLength(); j++) {
            Element cell = (Element)cells.item(j);
            cell.setAttribute("r", new Address(cell.getAttribute("r")).setRow(r).getA1());
        }

        return row;

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに文字列を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, String value) {

        if(value == null || value.isEmpty()) {
            setCellValueContent(address, "", "s");
        } else {
            int p = sharedStrings.add(value);
            setCellValueContent(address, String.valueOf(p), "s");
        }

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに真偽値を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, Boolean value) {

        setCellValueContent(address, (value ? "1" : "0"), "b");

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに日時を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, LocalDateTime value) {

        //https://poi.apache.org/apidocs/4.1/
        // Java LocalDateTime を Excel のシリアル値に変換し、文字列として挿入;
        double serial = getDateSerial(value) + getTimeSerial(value);
        setCellValueContent(address, String.valueOf(serial), null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに日付を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, LocalDate value) {

        //https://poi.apache.org/apidocs/4.1/
        // Java LocalDate を Excel のシリアル値に変換し、文字列として挿入;
        double serial = getDateSerial(value);
        setCellValueContent(address, String.valueOf(serial), null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに時刻を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, LocalTime value) {

        //https://poi.apache.org/apidocs/4.1/
        // Java LocalTime を Excel のシリアル値に変換し、文字列として挿入;
        double serial = getTimeSerial(value);
        setCellValueContent(address, String.valueOf(serial), null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに int 型の数値を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, int value) {

        setCellValueContent(address, String.valueOf(value), null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに long 型の整数を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, long value) {

        setCellValueContent(address, String.valueOf(value), null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに byte 型の整数を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, byte value) {

        setCellValueContent(address, String.valueOf(value), null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに short 型の整数を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, short value) {

        setCellValueContent(address, String.valueOf(value), null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに double 型の小数を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, double value) {

        setCellValueContent(address, String.valueOf(value), null);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 指定したアドレスのセルに float 型の小数を記入します。
     * @param address セルの A1 形式のアドレス
     * @param value 記入する値
     */
    public void setCellValue(String address, float value) {

        setCellValueContent(address, String.valueOf(value), null);

    }
    //------------------------------------------------------------------------------------------------
    private void setCellValueContent(String address, String value, String type) {

        Element cell = getCell(address, type);

        if(cell != null) {
            append(cell, "v", uri).setTextContent(value);
        }

    }
    // ------------------------------------------------------------------------------------------------
    private Element append(Element e, String name, String uri) {

        if(e.getElementsByTagName(name).getLength() == 0) {
            e.appendChild(xml.createElementNS(uri, name));
        }

        return (Element)e.getElementsByTagName(name).item(0);

    }
    // ------------------------------------------------------------------------------------------------
    private Element getCell(String address, String type) {

        Element cell = getCell(address);

        if(cell == null) {
            return null;
        }

        if(type == null || type.isEmpty()) {
            cell.removeAttribute("t");
        } else {
            cell.setAttribute("t", type);
        }

        return cell;

    }
    // ------------------------------------------------------------------------------------------------
    private Element getCell(String address) {

        int r = new Address(address).getRow();

        try {
            Element e = XPath.selectSingleNode(xml, "/x:worksheet/x:sheetData/x:row[@r=" + r + "]/x:c[@r='" + address + "']");

            // 該当のセルがあれば返す
            if(e != null) {
                return e;
            }

            Element row = getRow(r);
            if(row == null) {
                return null;
            }

            // 該当のセル行がなければ作成する
            e = xml.createElementNS(uri, "c");
            e.setAttribute("r", address);

            // 挿入箇所を探す
            Element where = null;

            NodeList cells = row.getElementsByTagName("c");
            for(int i = 0; i < cells.getLength(); i++) {
                Element cell = (Element)cells.item(i);
                if(cell.getAttribute("r").compareTo(address) > 0) {
                    where = cell;
                    break;
                }
            }

            if(where == null) {
                row.appendChild(e);
            } else {
                where.getParentNode().insertBefore(e, where);
            }

            return e;

        } catch(Exception ex) {
            return null;
        }

    }
    // ------------------------------------------------------------------------------------------------
    private Element getRow(int r) {

        Element e = XPath.selectSingleNode(xml, "/x:worksheet/x:sheetData/x:row[@r=" + r + "]");

        // 該当の行があれば返す
        if(e != null) {
            return e;
        }

        // 該当の行がなければ作成する
        e = xml.createElementNS(uri, "row");
        e.setAttribute("r", String.valueOf(r));

        // 挿入箇所を探す
        Element where = XPath.selectSingleNode(xml, "/x:worksheet/x:sheetData/x:row[@r>" + r + "][1]");
        Node sheetData = xml.getElementsByTagName("sheetData").item(0);

        if(where == null) {
            if(sheetData != null) {
                sheetData.appendChild(e);
                return e;
            }
        } else {
            where.getParentNode().insertBefore(e, where);
            return e;
        }

        return null;

    }
    //------------------------------------------------------------------------------------------------
    public void save() {

        save(this.file);

    }
    //------------------------------------------------------------------------------------------------
    private void save(File file) {

        try {
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            DOMSource source = new DOMSource(xml);
            StreamResult result = new StreamResult(file);
            transformer.transform(source, result);
        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    /**
     * XMLを文字列で返す（デバッグ用）
     */
    String getOuterXml() {

        return getOuterXml(null);

    }
    //------------------------------------------------------------------------------------------------
    private String getOuterXml(Node node) {

        // XMLを文字列で返す（デバッグ用）

        try {
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");

            DOMSource source = new DOMSource(node != null ? node : xml.getDocumentElement());
            StringWriter writer = new StringWriter();
            StreamResult result = new StreamResult(writer);
            transformer.transform(source, result);
            return writer.toString();

        } catch(Exception ex) {
            return ex.getMessage();
        }

    }
    //------------------------------------------------------------------------------------------------
    private double getDateSerial(LocalDateTime date) {

        return getDateSerial(date.toLocalDate());

    }
    //------------------------------------------------------------------------------------------------
    private double getDateSerial(LocalDate date) {

        LocalDate origin = LocalDate.of(1900, 1, 1);
        LocalDate ridge = LocalDate.of(1900, 2, 28);
        return (double)ChronoUnit.DAYS.between(origin, date) + (date.isAfter(ChronoLocalDate.from(ridge)) ? 2 : 1);

    }
    //------------------------------------------------------------------------------------------------
    private double getTimeSerial(LocalDateTime time) {

        return getTimeSerial(time.toLocalTime());

    }
    //------------------------------------------------------------------------------------------------
    private double getTimeSerial(LocalTime time) {

        double h = (double)time.getHour() / 24;
        double m = (double)time.getMinute() / (24 * 60);
        double s = (double)time.getSecond() / (24 * 60 * 60);
        return h + m + s;

    }
    //------------------------------------------------------------------------------------------------

}
//----------------------------------------------------------------------------------------------------







