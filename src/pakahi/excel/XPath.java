package pakahi.excel;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.util.ArrayList;
import java.util.List;



//----------------------------------------------------------------------------------------------------
/**
 * ノード選択のための XPath 操作のクラスです。
 */
public class XPath {

    //------------------------------------------------------------------------------------------------
    /**
     * 複数ノードを選択します。
     * @param xml DOMドキュメント
     * @param path XPath パターン
     * @return 要素のリスト
     */
    public static List<Element> selectNodes(Document xml, String path) {

        List<Element> items = new ArrayList<>();

        try {
            javax.xml.xpath.XPath xpath = XPathFactory.newInstance().newXPath();
            xpath.setNamespaceContext(new ExcelNamespaceContext());
            NodeList nodes = (NodeList)xpath.compile(path).evaluate(xml, XPathConstants.NODESET);

            if(nodes == null || nodes.getLength() == 0) {
                return items;
            }

            for(int i = 0; i < nodes.getLength(); i++) {
                items.add((Element)nodes.item(i));
            }

            return items;

        } catch(Exception ex) {
            return null;
        }

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 単一ノードを選択します。
     * @param xml DOMドキュメント
     * @param path XPath パターン
     * @return 要素。見つからなければ null が返されます。
     */
    public static Element selectSingleNode(Document xml, String path) {

        try {
            javax.xml.xpath.XPath xpath = XPathFactory.newInstance().newXPath();
            xpath.setNamespaceContext(new ExcelNamespaceContext());
            NodeList nodes = (NodeList)xpath.compile(path).evaluate(xml, XPathConstants.NODESET);

            if(nodes == null || nodes.getLength() == 0) {
                return null;
            }

            return (Element)nodes.item(0);

        } catch(Exception ex) {
            return null;
        }

    }
    //------------------------------------------------------------------------------------------------

}
//----------------------------------------------------------------------------------------------------
