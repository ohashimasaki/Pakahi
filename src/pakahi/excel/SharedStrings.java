package pakahi.excel;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;
import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.StringReader;
import java.io.StringWriter;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;



//----------------------------------------------------------------------------------------------------
/**
 * 共有文字列（/xl/sharedString.xml）を保持・操作するクラスです。
 */
public class SharedStrings {

    Path template;
    File file;
    Document xml;
    List<String> sharedStrings = new ArrayList<>();
    int count = 0;
    int uniqueCount = 0;
    NamespaceContext context;


    //------------------------------------------------------------------------------------------------
    /**
     * 共有文字列（/xl/sharedString.xml）を保持・操作するクラスのコンストラクタです。
     * @param template テンプレートとなる SpreadsheetML 構造のパス
     */
    public SharedStrings(Path template) {

        this.template = template;
        file = Paths.get(template.toString(), "xl/sharedStrings.xml").toFile();

        context = new ExcelNamespaceContext();

        if(load()) {
            List<Element> nodes = XPath.selectNodes(xml, "/x:sst/x:si");
            if(nodes != null) {
                count = nodes.size();
            }

            List<Element> strings = XPath.selectNodes(xml, "/x:sst/x:si/x:t[1]|/x:sst/x:si/x:r[1]/x:t[1]");
            if(strings != null) {
                uniqueCount = strings.size();
                sharedStrings.addAll(strings.stream().map(Element::getTextContent).collect(Collectors.toList()));
            }
        }

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 共有文字列に文字列を追加します。
     * @param s 追加する文字列
     */
    public int add(String s) {

        if(sharedStrings.contains(s)) {
            return sharedStrings.indexOf(s);
        }

        Element si = xml.createElement("si");
        Element t = xml.createElement("t");
        t.setTextContent(s);
        si.appendChild(t);
        xml.getDocumentElement().appendChild(si);

        sharedStrings.add(s);

        count++;
        uniqueCount++;

        return sharedStrings.indexOf(s);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 共有文字列を一時フォルダ内に保存します。
     */
    public void save() {

        save(file);

    }
    //------------------------------------------------------------------------------------------------
    private void save(File file) {

        try {
            Element d = xml.getDocumentElement();
            d.setAttribute("count", String.valueOf(count));
            d.setAttribute("uniqueCount", String.valueOf(uniqueCount));
            save(xml, file);

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    private void save(Document xml, File file) {

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
    private boolean load() {

        if( ! file.exists()) {
            return restore();
        }

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            xml = factory.newDocumentBuilder().parse(file);

            if(xml.getDocumentElement().hasAttribute("count")) {
                count = Integer.parseInt(xml.getDocumentElement().getAttribute("count"));
            }
            if(xml.getDocumentElement().hasAttribute("uniqueCount")) {
                uniqueCount = Integer.parseInt(xml.getDocumentElement().getAttribute("uniqueCount"));
            }

            return true;

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
            return restore();
        }

    }
    //------------------------------------------------------------------------------------------------
    private boolean restore() {

        String uri = context.getNamespaceURI("x");

        String t = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
        t += "<sst uniqueCount=\"0\" count=\"0\" xmlns=\"" + uri + "\"></sst>";

        InputSource s = new InputSource(new StringReader(t));

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            xml = factory.newDocumentBuilder().parse(s);
            save();
            addWorkbookRel();
            addContentType();
            return true;

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
            return false;
        }

    }
    //------------------------------------------------------------------------------------------------
    private void addWorkbookRel() throws FileNotFoundException {

        File file = Paths.get(template.toString(), "xl/_rels/workbook.xml.rels").toFile();

        if( ! file.exists()) {
            throw new FileNotFoundException("File \"xl/_rels/workbook.xml.rels\" not found.");
        }

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            Document xml = factory.newDocumentBuilder().parse(file);

            if(XPath.selectSingleNode(xml, "/r:Relationships/r:Relationship[@Target='sharedStrings.xml']") != null) {
                return;
            }

            List<Element> items = XPath.selectNodes(xml, "/r:Relationships/r:Relationship");

            int c = items == null ? 0 : Collections.max(items.stream()
                .map(e -> Integer.parseInt(e.getAttribute("Id").replace("rId", "")))
                .collect(Collectors.toList())) + 1;

            Element e = xml.createElementNS(context.getNamespaceURI("r"), "Relationship");
            e.setAttribute("Id", "rId" + c);
            e.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            e.setAttribute("Target", "sharedStrings.xml");
            xml.getDocumentElement().appendChild(e);
            save(xml, file);

        } catch(Exception ex) {
            System.out.println(Arrays.toString(ex.getStackTrace()));
        }

    }
    //------------------------------------------------------------------------------------------------
    private void addContentType() throws FileNotFoundException {

        File file = Paths.get(template.toString(), "[Content_Types].xml").toFile();

        if( ! file.exists()) {
            throw new FileNotFoundException("File \"[Content_Types].xml\" not found.");
        }

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            Document xml = factory.newDocumentBuilder().parse(file);

            if(XPath.selectSingleNode(xml, "/t:Types/t:Override[@PartName='/xl/sharedStrings.xml']") != null) {
                return;
            }

            Element e = xml.createElementNS(context.getNamespaceURI("t"), "Override");
            e.setAttribute("PartName", "/xl/sharedStrings.xml");
            e.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
            xml.getDocumentElement().appendChild(e);
            save(xml, file);

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
//        transformer.setOutputProperty("omit-xml-declaration", "yes");
        transformer.setOutputProperty("indent", "yes");

            StringWriter writer = new StringWriter();
            transformer.transform(new DOMSource(node != null ? node : xml.getDocumentElement()), new StreamResult(writer));
            return writer.toString();

        } catch(Exception ex) {
            return ex.getMessage();
        }

    }
    //------------------------------------------------------------------------------------------------

}
//----------------------------------------------------------------------------------------------------
