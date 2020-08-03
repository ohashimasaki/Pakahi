package pakahi.excel;

import javax.xml.namespace.NamespaceContext;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;



//----------------------------------------------------------------------------------------------------
/**
 * 名前空間 URI を管理するクラスです。
 */
public class ExcelNamespaceContext implements NamespaceContext {

    private final HashMap<String, String> namespaces = new HashMap<>() {{
        // DEFAULT NAMESPACE URI
        put("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        // OTHER NAMESPACE URIs (ADD IF NEEDED)
        put("o", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        put("t", "http://schemas.openxmlformats.org/package/2006/content-types");
        put("r", "http://schemas.openxmlformats.org/package/2006/relationships");
    }};

    //------------------------------------------------------------------------------------------------
    /**
     * プレフィクスから名前空間 URI を返します。
     * @param prefix プレフィクス
     * @return URI
     */
    public String getNamespaceURI(String prefix) {

        return namespaces.get(prefix);

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 名前空間 URI からプレフィクスを返します。
     * @param uri URI
     * @return プレフィクス。見つからなければ null が返されます。
     */
    public String getPrefix(String uri) {

        if(uri == null || uri.isEmpty() || ! namespaces.containsValue(uri)) {
            return null;
        }

        for(String prefix : namespaces.keySet()) {
            if(namespaces.get(prefix).equals(uri)) {
                return prefix;
            }
        }

        return null;

    }
    //------------------------------------------------------------------------------------------------
    /**
     * 名前空間 URI からプレフィクスのリストを返します。
     * @param uri URI
     * @return プレフィクスのリスト
     */
    public Iterator getPrefixes(String uri) {

        return new ArrayList<>(namespaces.keySet()).iterator();

    }
    //------------------------------------------------------------------------------------------------

}
//----------------------------------------------------------------------------------------------------