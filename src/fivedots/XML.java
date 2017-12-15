package fivedots;

/*  XML.java
    Andrew Davison, ad@fivedots.coe.psu.ac.th, December 2016
    Heavily edited and formated by K. Kacprzak 2017
    XML utilities:
    * XML Document IO
    * DOM data extraction
    * String extraction
    * XLS transforming
    * Flat XML filter selection
    Useful code:
    http://www.drdobbs.com/jvm/easy-dom-parsing-in-java/231002580
    http://www.java2s.com/Code/Java/XML/FindAllElementsByTagName.htm
    XSLT tutorial at W3Schools
    http://www.w3schools.com/xml/xsl_intro.asp
    Also:
    "Appendix B. The XSLT You Need for OpenOffice.org"
    http://books.evc-cit.info/apb.php
*/

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.StringReader;
import java.io.StringWriter;
import java.net.URL;
import java.util.ArrayList;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.Source;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

public class XML {

    private static final String INDENT_FNM = "indent.xsl";
    public static Document loadDoc(String fnm) 
            throws SAXException, ParserConfigurationException, IOException {
        Document doc = null;
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            dbf.setValidating(false);
            dbf.setNamespaceAware(true);
            DocumentBuilder db = dbf.newDocumentBuilder();
            doc = db.parse(new FileInputStream(new File(fnm)));
        } catch (ParserConfigurationException | SAXException | IOException e) {
            System.out.println(e);
        }
        return doc;
    } 

    public static Document url2Doc(String urlStr) {
        Document doc = null;
        try {
            URL xmlUrl = new URL(urlStr);
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            dbf.setValidating(false);
            dbf.setNamespaceAware(true);
            DocumentBuilder db = dbf.newDocumentBuilder();
            doc = db.parse(xmlUrl.openStream());
        } catch (ParserConfigurationException | IOException | SAXException e) {
            System.out.println(e);
        }
        return doc;
    }

    public static Document str2Doc(String xmlStr) {
        Document doc = null;
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            dbf.setValidating(false);
            dbf.setNamespaceAware(true);  // false);
            DocumentBuilder db = dbf.newDocumentBuilder();
            doc = db.parse(new InputSource(new StringReader(xmlStr)));
        } catch (ParserConfigurationException | SAXException | IOException e) {
            System.out.println(e);
        }
        return doc;
    }

    public static void saveDoc(Document doc, String xmlFnm) 
            throws TransformerConfigurationException, TransformerException {
        try {
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer t = tf.newTransformer();
            t.setOutputProperty(OutputKeys.INDENT, "yes");
            t.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
            DOMSource source = new DOMSource(doc);
            StreamResult result = new StreamResult(new File(xmlFnm));
            t.transform(source, result);
            System.out.println("Saved document to " + xmlFnm);
        } catch (IllegalArgumentException | TransformerException e) {
            System.out.println("Unable to save document to " + xmlFnm);
            System.out.println("  " + e);
        }
    }

    public static Node getNode(String tagName, NodeList nodes) {
        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            if (node.getNodeName().equalsIgnoreCase(tagName)) {
                return node;
            }
        }
        return null;
    }

    public static String getNodeValue(String tagName, NodeList nodes) {
        if (nodes == null) return "";
        for (int i = 0; i < nodes.getLength(); i++) {
            Node n = nodes.item(i);
            if (n.getNodeName().equalsIgnoreCase(tagName)) return getNodeValue(n);
        }
        return "";
    }

    public static String getNodeValue(Node node) {
        if (node == null) return "";
        NodeList childNodes = node.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node n = childNodes.item(i);
            if (n.getNodeType() == Node.TEXT_NODE) return n.getNodeValue().trim();
        }
        return "";
    }  

    public static ArrayList<String> getNodeValues(NodeList nodes) {
        if (nodes == null) return null;
        ArrayList<String> vals = new ArrayList<>();
        for (int i = 0; i < nodes.getLength(); i++) {
            String val = getNodeValue(nodes.item(i));
            if (val != null) vals.add(val);
        }
        return vals;
    } 

    public static String getNodeAttr(String attrName, Node node) {
        if (node == null) return "";
        NamedNodeMap attrs = node.getAttributes();
        if (attrs == null) return "";
        for (int i = 0; i < attrs.getLength(); i++) {
            Node attr = attrs.item(i);
            if (attr.getNodeName().equalsIgnoreCase(attrName)) 
                return attr.getNodeValue().trim();
        }
        return "";
    }

    public static Object[][] getAllNodeValues(NodeList rowNodes, String[] colIDs) {
        int numRows = rowNodes.getLength();
        int numCols = colIDs.length;
        Object[][] data = new Object[numRows + 1][numCols];
        for (int col = 0; col < numCols; col++) {
            data[0][col] = Lo.capitalize(colIDs[col]);
        }
        for (int i = 0; i < numRows; i++) {
            NodeList colNodes = rowNodes.item(i).getChildNodes();
            for (int col = 0; col < numCols; col++) {
                data[i + 1][col] = getNodeValue(colIDs[col], colNodes);
            }
        }
        return data;
    }

    public static String getDocString(Document doc) {
        try {
            TransformerFactory trf = TransformerFactory.newInstance();
            Transformer tr = trf.newTransformer();
            DOMSource domSource = new DOMSource(doc);
            StringWriter sw = new StringWriter();
            StreamResult result = new StreamResult(sw);
            tr.transform(domSource, result);
            System.out.println("Extracting string from document");
            String xmlStr = sw.toString();
            return indent2Str(xmlStr);
        } catch (TransformerException ex) {
            System.out.println("Could not convert document to a string");
            return null;
        }
    }  

    public static String applyXSLT(String xmlFnm, String xslFnm) {
        try {
            TransformerFactory tf = TransformerFactory.newInstance();
            Source xslt = new StreamSource(new File(xslFnm));
            Transformer t = tf.newTransformer(xslt);
            System.out.println("Applying filter " + xslFnm + " to " + xmlFnm);
            Source text = new StreamSource(new File(xmlFnm));
            StreamResult result = new StreamResult(new StringWriter());
            t.transform(text, result);
            return result.getWriter().toString();
        } catch (TransformerException e) {
            System.out.println("Unable to transform " + xmlFnm + " with " + xslFnm);
            System.out.println("  " + e);
            return null;
        }
    } 

    public static String applyXSLT2str(String xmlStr, String xslFnm) {
        try {
            TransformerFactory tf = TransformerFactory.newInstance();
            Source xslt = new StreamSource(new File(xslFnm));
            Transformer t = tf.newTransformer(xslt);
            System.out.println("Applying the filter in " + xslFnm);
            Source text = new StreamSource(new StringReader(xmlStr));
            StreamResult result = new StreamResult(new StringWriter());
            t.transform(text, result);
            return result.getWriter().toString();
        } catch (TransformerException e) {
            System.out.println("Unable to transform the string");
            System.out.println("  " + e);
            return null;
        }
    }

    public static String indent(String xmlFnm) {
        return applyXSLT(xmlFnm, FileIO.getUtilsFolder() + INDENT_FNM);
    }

    public static String indent2Str(String xmlStr) {
        return applyXSLT2str(xmlStr, FileIO.getUtilsFolder() + INDENT_FNM);
    }

    public static String getFlatFilterName(String docType) {
        if (null != docType) switch (docType) {
            case Lo.WRITER_STR:
                return "OpenDocument Text Flat XML";
            case Lo.CALC_STR:
                return "OpenDocument Spreadsheet Flat XML";
            case Lo.DRAW_STR:
                return "OpenDocument Drawing Flat XML";
            case Lo.IMPRESS_STR:
                return "OpenDocument Presentation Flat XML";
        }
        System.out.println("No Flat XML filter for this document type; using Flat text");
        return "OpenDocument Text Flat XML";
    }

}  // end of XML class
