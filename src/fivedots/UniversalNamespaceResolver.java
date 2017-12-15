package fivedots;

import java.util.Iterator;
import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import org.w3c.dom.Document;

public class UniversalNamespaceResolver implements NamespaceContext {

    private final Document sourceDoc;
    public UniversalNamespaceResolver(Document document) {
        sourceDoc = document;
    }

    @Override
    public String getNamespaceURI(String prefix) {
        if (prefix.equals(XMLConstants.DEFAULT_NS_PREFIX)) {
            return sourceDoc.lookupNamespaceURI(null);
        } else {
            return sourceDoc.lookupNamespaceURI(prefix);
        }
    }

    @Override
    public String getPrefix(String namespaceURI) {
        return sourceDoc.lookupPrefix(namespaceURI);
    }

    @Override
    public Iterator getPrefixes(String namespaceURI) {
        return null;
    }

}