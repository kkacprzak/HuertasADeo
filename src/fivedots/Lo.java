package fivedots;

/*  Lo.java
    Andrew Davison, ad@fivedots.coe.psu.ac.th, February 2015
    Heavily edited and formated by K. Kacprzak 2017
    A growing collection of utility functions to make Office
    easier to use. They are currently divided into the following groups:
    * interface object creation (uses generics)
    * office starting
    * office shutdown
    * document opening
    * document creation
    * document saving
    * document closing
    * initialization via Addon-supplied context
    * initialization via script context
    * dispatch
    * UNO cmds
    * use Inspectors extension
    * color methods
    * other utils
    * container manipulation
 */

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XIntrospection;
import com.sun.star.beans.XIntrospectionAccess;
import com.sun.star.beans.XPropertySet;
import java.net.URLClassLoader;
import com.sun.star.bridge.XBridge;
import com.sun.star.bridge.XBridgeFactory;
import com.sun.star.comp.beans.OOoBean;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.connection.XConnection;
import com.sun.star.connection.XConnector;
import com.sun.star.container.XChild;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNamed;
import com.sun.star.document.MacroExecMode;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XStorable;
import com.sun.star.uno.Exception;
import com.sun.star.io.IOException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.reflection.XIdlMethod;
import com.sun.star.script.provider.XScriptContext;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.Arrays;

public class Lo {

    public static final int UNKNOWN = 0;
    public static final int WRITER = 1;
    public static final int BASE = 2;
    public static final int CALC = 3;
    public static final int DRAW = 4;
    public static final int IMPRESS = 5;
    public static final int MATH = 6;
    public static final String UNKNOWN_STR = "unknown";
    public static final String WRITER_STR = "swriter";
    public static final String BASE_STR = "sbase";
    public static final String CALC_STR = "scalc";
    public static final String DRAW_STR = "sdraw";
    public static final String IMPRESS_STR = "simpress";
    public static final String MATH_STR = "smath";
    public static final String UNKNOWN_SERVICE = "com.sun.frame.XModel";
    public static final String WRITER_SERVICE = "com.sun.star.text.TextDocument";
    public static final String BASE_SERVICE = "com.sun.star.sdb.OfficeDatabaseDocument";
    public static final String CALC_SERVICE = "com.sun.star.sheet.SpreadsheetDocument";
    public static final String DRAW_SERVICE = "com.sun.star.drawing.DrawingDocument";
    public static final String IMPRESS_SERVICE = "com.sun.star.presentation.PresentationDocument";
    public static final String MATH_SERVICE = "com.sun.star.formula.FormulaProperties";
    private static final int SOCKET_PORT = 8100;
    public static final String WRITER_CLSID = "8BC6B165-B1B2-4EDD-aa47-dae2ee689dd6";
    public static final String CALC_CLSID = "47BBB4CB-CE4C-4E80-a591-42d9ae74950f";
    public static final String DRAW_CLSID = "4BAB8970-8A3B-45B3-991c-cbeeac6bd5e3";
    public static final String IMPRESS_CLSID = "9176E48A-637A-4D1F-803b-99d9bfac1047";
    public static final String MATH_CLSID = "078B7ABA-54FC-457F-8551-6147e776a997";
    public static final String CHART_CLSID = "12DCAE26-281F-416F-a234-c3086127382e";
    private static XComponentContext xcc = null;
    private static XDesktop xDesktop = null;
    private static XMultiComponentFactory mcFactory = null;
    private static XMultiServiceFactory msFactory = null;
    private static XComponent bridgeComponent = null;
    private static boolean isOfficeTerminated = false;

    public static XComponentContext getContext() {
        return xcc;
    }

    public static XDesktop getDesktop() {
        return xDesktop;
    }

    public static XMultiComponentFactory getComponentFactory() {
        return mcFactory;
    }

    public static XMultiServiceFactory getServiceFactory() {
        return msFactory;
    }

    public static XComponent getBridge() {
        return bridgeComponent;
    }

    public static void setOOoBean(OOoBean oob) {
        try {
            xcc = oob.getOOoConnection().getComponentContext();
            if (xcc == null) {
                System.out.println("No component context found in OOoBean");
            } else {
                mcFactory = xcc.getServiceManager();
            }
            xDesktop = oob.getOOoDesktop();
            msFactory = oob.getMultiServiceFactory();
        } catch (java.lang.Exception e) {
            System.out.println("Couldn't initialize LO using OOoBean: " + e);
        }
    }

    public static <T> T qi(Class<T> aType, Object o) {
        return UnoRuntime.queryInterface(aType, o);
    }

    public static <T> T createInstanceMSF(Class<T> aType, String serviceName) {
        if (msFactory == null) {
            System.out.println("No document found");
            return null;
        }
        T interfaceObj = null;
        try {
            Object o = msFactory.createInstance(serviceName);
            interfaceObj = Lo.qi(aType, o);
        } catch (Exception e) {
            System.out.println("Couldn't create interface for \"" + serviceName + "\": " + e);
        }
        return interfaceObj;
    }

    public static <T> T createInstanceMSF(Class<T> aType, String serviceName,
            XMultiServiceFactory msf) {
        if (msf == null) {
            System.out.println("No document found");
            return null;
        }
        T interfaceObj = null;
        try {
            Object o = msf.createInstance(serviceName);
            interfaceObj = Lo.qi(aType, o);
        } catch (Exception e) {
            System.out.println("Couldn't create interface for \"" + serviceName + "\":\n  " + e);
        }
        return interfaceObj;
    }

    public static <T> T createInstanceMCF(Class<T> aType, String serviceName) {
        if ((xcc == null) || (mcFactory == null)) {
            System.out.println("No office connection found");
            return null;
        }
        T interfaceObj = null;
        try {
            Object o = mcFactory.createInstanceWithContext(serviceName, xcc);
            interfaceObj = Lo.qi(aType, o);
        } catch (Exception e) {
            System.out.println("Couldn't create interface for \"" + serviceName + "\": " + e);
        }
        return interfaceObj;
    }

    public static <T> T createInstanceMCF(Class<T> aType, String serviceName, Object[] args) {
        if ((xcc == null) || (mcFactory == null)) {
            System.out.println("No office connection found");
            return null;
        }
        T interfaceObj = null;
        try {
            Object o = mcFactory.createInstanceWithArgumentsAndContext(serviceName, args, xcc);
            interfaceObj = Lo.qi(aType, o);
        } catch (Exception e) {
            System.out.println("Couldn't create interface for \"" + serviceName + "\": " + e);
        }
        return interfaceObj;
    }

    public static <T> T getParent(Object aComponent, Class<T> aType) {
        XChild xAsChild = Lo.qi(XChild.class, aComponent);
        return Lo.qi(aType, xAsChild.getParent());
    }

    public static XComponentLoader loadOffice() {
        return loadOffice(true);
    }

    public static XComponentLoader loadSocketOffice() {
        return loadOffice(false);
    }

    public static XComponentLoader loadOffice(boolean usingPipes) {
        System.out.println("Loading Office...");
        if (usingPipes) {
            xcc = bootstrapContext();
        } else {
            xcc = socketContext();
        }
        if (xcc == null) {
            System.out.println("Office context could not be created");
            System.exit(1);
        }
        mcFactory = xcc.getServiceManager();
        if (mcFactory == null) {
            System.out.println("Office Service Manager is unavailable");
            System.exit(1);
        }
        xDesktop = createInstanceMCF(XDesktop.class, "com.sun.star.frame.Desktop");
        if (xDesktop == null) {
            System.out.println("Could not create a desktop service");
            System.exit(1);
        }
        return Lo.qi(XComponentLoader.class, xDesktop);
    }

    private static XComponentContext bootstrapContext() {
        XComponentContext bcxcc = null;
        try {
            bcxcc = Bootstrap.bootstrap();
        } catch (BootstrapException e) {
            System.out.println("Unable to bootstrap Office "+ e);
        }
        return bcxcc;
    }

    private static XComponentContext socketContext() {
        XComponentContext scxcc = null;
        try {
            String[] cmdArray = new String[3];
            cmdArray[0] = "soffice";
            cmdArray[1] = "-headless";
            cmdArray[2] = "-accept=socket,host=localhost,port="
                    + SOCKET_PORT + ";urp;";
            Process p = Runtime.getRuntime().exec(cmdArray);
            if (p != null) {
                System.out.println("Office process created");
            }
            delay(5000);
            XComponentContext localContext
                    = Bootstrap.createInitialComponentContext(null);
            XMultiComponentFactory localFactory = localContext.getServiceManager();
            XConnector connector = Lo.qi(XConnector.class,
                    localFactory.createInstanceWithContext(
                            "com.sun.star.connection.Connector", localContext));
            XConnection connection = connector.connect(
                    "socket,host=localhost,port=" + SOCKET_PORT);
            XBridgeFactory bridgeFactory = Lo.qi(XBridgeFactory.class,
                    localFactory.createInstanceWithContext(
                            "com.sun.star.bridge.BridgeFactory", localContext));
            XBridge bridge = bridgeFactory.createBridge("socketBridgeAD", "urp", connection, null);
            bridgeComponent = Lo.qi(XComponent.class, bridge);
            XMultiComponentFactory serviceManager = Lo.qi(XMultiComponentFactory.class,
                    bridge.getInstance("StarOffice.ServiceManager"));
            XPropertySet props = Lo.qi(XPropertySet.class, serviceManager);
            Object defaultContext = props.getPropertyValue("DefaultContext");
            scxcc = Lo.qi(XComponentContext.class, defaultContext);
        } catch (java.lang.Exception e) {
            System.out.println("Unable to socket connect to Office " + e);
        }
        return scxcc;
    }

    public static void closeOffice() {
        System.out.println("Closing Office");
        if (xDesktop == null) {
            System.out.println("No office connection found");
            return;
        }
        if (isOfficeTerminated) {
            System.out.println("Office has already been requested to terminate");
            return;
        }
        int numTries = 1;
        while (!isOfficeTerminated && (numTries < 4)) {
            delay(200);
            isOfficeTerminated = tryToTerminate(numTries);
            numTries++;
        }
    }

    public static boolean tryToTerminate(int numTries) {
        try {
            boolean isDead = xDesktop.terminate();
            if (isDead) {
                if (numTries > 1) {
                    System.out.println(numTries + ". Office terminated");
                } else {
                    System.out.println("Office terminated");
                }
            } else {
                System.out.println(numTries + ". Office failed to terminate");
            }
            return isDead;
        } catch (com.sun.star.lang.DisposedException e) {
            System.out.println("Office link disposed");
            return true;
        } catch (java.lang.Exception e) {
            System.out.println("Termination exception: " + e);
            return false;
        }
    }

    public static void killOffice() {
        try {
            Runtime.getRuntime().exec("cmd /c lokill.bat");
            System.out.println("Killed Office");
        } catch (java.lang.Exception e) {
            System.out.println("Unable to kill Office: " + e);
        }
    }

    public static XComponent openFlatDoc(
            String fnm, String docType, XComponentLoader loader) {
        String nm = XML.getFlatFilterName(docType);
        System.out.println("Flat filter Name: " + nm);
        return openDoc(fnm, loader, Props.makeProps("FilterName", nm));
    }

    public static XComponent openDoc(String fnm, XComponentLoader loader) {
        return openDoc(fnm, loader, Props.makeProps("Hidden", true));
    }

    public static XComponent openReadOnlyDoc(String fnm, XComponentLoader loader) {
        return openDoc(fnm, loader, Props.makeProps("Hidden", true, "ReadOnly", true));
    }

    public static XComponent openDoc(
            String fnm,
            XComponentLoader loader,
            PropertyValue[] props) {
        if (fnm == null) {
            System.out.println("Filename is null");
            return null;
        }
        String openFileURL;
        if (!FileIO.isOpenable(fnm)) {
            if (isURL(fnm)) {
                System.out.println("Will treat filename as a URL: \"" + fnm + "\"");
                openFileURL = fnm;
            } else {
                return null;
            }
        } else {
            System.out.println("Opening " + fnm);
            openFileURL = FileIO.fnmToURL(fnm);
            if (openFileURL == null) {
                return null;
            }
        }
        XComponent doc = null;
        try {
            doc = loader.loadComponentFromURL(openFileURL, "_blank", 0, props);
            msFactory = Lo.qi(XMultiServiceFactory.class, doc);
        } catch (Exception e) {
            System.out.println("Unable to open the document " + e);
        }
        return doc;
    }

    public static boolean isURL(String fnm) {
        try {
            java.net.URL u = new java.net.URL(fnm);
            u.toURI();
            return true;
        } catch (java.net.MalformedURLException | java.net.URISyntaxException e) {
            return false;
        }
    }

    public static String ext2DocType(String ext) {
        switch (ext) {
            case "odt":
                return WRITER_STR;
            case "odp":
                return IMPRESS_STR;
            case "odg":
                return DRAW_STR;
            case "ods":
                return CALC_STR;
            case "odb":
                return BASE_STR;
            case "odf":
                return MATH_STR;
            default:
                System.out.println("Do not recognize extension \"" + ext + "\"; using writer");
                return WRITER_STR;
        }
    }

    public static String docTypeStr(int docTypeVal) {
        switch (docTypeVal) {
            case WRITER:
                return WRITER_STR;
            case IMPRESS:
                return IMPRESS_STR;
            case DRAW:
                return DRAW_STR;
            case CALC:
                return CALC_STR;
            case BASE:
                return BASE_STR;
            case MATH:
                return MATH_STR;
            default:
                System.out.println("Do not recognize extension \"" + docTypeVal + "\"; using writer");
                return WRITER_STR;
        }
    }

    public static XComponent createDoc(String docType, XComponentLoader loader) {
        return createDoc(docType, loader, Props.makeProps("Hidden", true));
    }

    public static XComponent createMacroDoc(String docType, XComponentLoader loader) {
        return createDoc(docType, loader, Props.makeProps("Hidden", false,
                "MacroExecutionMode", MacroExecMode.ALWAYS_EXECUTE_NO_WARN));
    }

    public static XComponent createDoc(
            String docType,
            XComponentLoader loader,
            PropertyValue[] props) {
        System.out.println("Creating Office document " + docType);
        XComponent doc = null;
        try {
            doc = loader.loadComponentFromURL("private:factory/" + docType, "_blank", 0, props);
            msFactory = Lo.qi(XMultiServiceFactory.class, doc);
        } catch (Exception e) {
            System.out.println("Could not create a document");
        }
        return doc;
    }

    public static XComponent createDocFromTemplate(
            String templatePath,
            XComponentLoader loader) {
        if (!FileIO.isOpenable(templatePath)) {
            return null;
        }
        System.out.println("Opening template " + templatePath);
        String templateURL = FileIO.fnmToURL(templatePath);
        if (templateURL == null) {
            return null;
        }
        PropertyValue[] props = Props.makeProps(
                "Hidden",
                true,
                "AsTemplate",
                true);
        XComponent doc = null;
        try {
            doc = loader.loadComponentFromURL(templateURL, "_blank", 0, props);
            msFactory = Lo.qi(XMultiServiceFactory.class, doc);
        } catch (Exception e) {
            System.out.println("Could not create document from template: " + e);
        }
        return doc;
    }

    public static void save(Object odoc) {
        XStorable store = Lo.qi(XStorable.class, odoc);
        try {
            store.store();
            System.out.println("Saved the document by overwriting");
        } catch (IOException e) {
            System.out.println("Could not save the document");
        }
    }

    public static void saveDoc(Object odoc, String fnm) {
        XStorable store = Lo.qi(XStorable.class, odoc);
        XComponent doc = Lo.qi(XComponent.class, odoc);
        int docType = Info.reportDocType(doc);
        storeDoc(store, docType, fnm, null);
    }

    public static void saveDoc(Object odoc, String fnm, String password) {
        XStorable store = Lo.qi(XStorable.class, odoc);
        XComponent doc = Lo.qi(XComponent.class, odoc);
        int docType = Info.reportDocType(doc);
        storeDoc(store, docType, fnm, password);
    }

    public static void saveDoc(Object odoc, String fnm, String format, String password) {
        XStorable store = Lo.qi(XStorable.class, odoc);
        storeDocFormat(store, fnm, format, password);
    }

    public static void storeDoc(XStorable store, int docType, String fnm, String password) {
        String ext = Info.getExt(fnm);
        String format = "Text";
        if (ext == null) System.out.println("Assuming a text format");
        else format = ext2Format(docType, ext);
        storeDocFormat(store, fnm, format, password);
    }

    public static String ext2Format(String ext) {
        return ext2Format(Lo.UNKNOWN, ext);
    }

    public static String ext2Format(int docType, String ext) {
        switch (ext) {
            case "doc":
                return "MS Word 97";
            case "docx":
                return "Office Open XML Text";
            case "rtf":
                if (docType == Lo.CALC) return "Rich Text Format (StarCalc)";
                else return "Rich Text Format";
            case "odt":
                return "writer8";
            case "ott":
                return "writer8_template";
            case "pdf":
                switch (docType) {
                    case Lo.WRITER:
                        return "writer_pdf_Export";
                    case Lo.IMPRESS:
                        return "impress_pdf_Export";
                    case Lo.DRAW:
                        return "draw_pdf_Export";
                    case Lo.CALC:
                        return "calc_pdf_Export";
                    case Lo.MATH:
                        return "math_pdf_Export";
                    default:
                        return "writer_pdf_Export";
                }
            case "txt":
                return "Text";
            case "ppt":
                return "MS PowerPoint 97";
            case "pptx":
                return "Impress MS PowerPoint 2007 XML";
            case "odp":
                return "impress8";
            case "odg":
                return "draw8";
            case "jpg":
                if (docType == Lo.IMPRESS) return "impress_jpg_Export";
                else return "draw_jpg_Export";
            case "png":
                if (docType == Lo.IMPRESS) return "impress_png_Export";
                else return "draw_png_Export";
            case "xls":
                return "MS Excel 97";
            case "xlsx":
                return "Calc MS Excel 2007 XML";
            case "csv":
                return "Text - txt - csv (StarCalc)";
            case "ods":
                return "calc8";
            case "odb":
                return "StarOffice XML (Base)";
            case "htm":
            case "html":
                switch (docType) {
                    case Lo.WRITER:
                        return "HTML (StarWriter)"; 
                    case Lo.IMPRESS:
                        return "impress_html_Export";
                    case Lo.DRAW:
                        return "draw_html_Export";
                    case Lo.CALC:
                        return "HTML (StarCalc)";
                    default:
                        return "HTML";
                }
            case "xhtml":
                switch (docType) {
                    case Lo.WRITER:
                        return "XHTML Writer File";
                    case Lo.IMPRESS:
                        return "XHTML Impress File";
                    case Lo.DRAW:
                        return "XHTML Draw File";
                    case Lo.CALC:
                        return "XHTML Calc File";
                    default:
                        return "XHTML Writer File"; 
                }
            case "xml":
                switch (docType) {
                    case Lo.WRITER:
                        return "OpenDocument Text Flat XML";
                    case Lo.IMPRESS:
                        return "OpenDocument Presentation Flat XML";
                    case Lo.DRAW:
                        return "OpenDocument Drawing Flat XML";
                    case Lo.CALC:
                        return "OpenDocument Spreadsheet Flat XML";
                    default:
                        return "OpenDocument Text Flat XML"; 
        }
            default: 
                System.out.println("Do not recognize extension \"" + ext + "\"; using text");
                return "Text";
        }
    } 

    public static void storeDocFormat(
            XStorable store, String fnm, 
            String format, String password) {
        System.out.println("Saving the document in " + fnm);
        System.out.println("Using format: " + format);
        try {
            String saveFileURL = FileIO.fnmToURL(fnm);
            if (saveFileURL == null) return;
            PropertyValue[] storeProps;
            if (password == null) storeProps = Props.makeProps("Overwrite", true, "FilterName", format);
            else {
                String[] nms = new String[]{"Overwrite", "FilterName", "Password"};
                Object[] vals = new Object[]{true, format, password};
                storeProps = Props.makeProps(nms, vals);
            }
            store.storeToURL(saveFileURL, storeProps);
        } catch (IOException e) {
            System.out.println("Could not save " + fnm + ": " + e);
        }
    }

    public static void closeDoc(Object doc) {
        try {
            XCloseable closeable = Lo.qi(XCloseable.class, doc);
            close(closeable);
        } catch (com.sun.star.lang.DisposedException e) {
            System.out.println("Document close failed since Office link disposed "+e);
        }
    }

    public static void close(XCloseable closeable) {
        if (closeable == null) return;
        System.out.println("Closing the document");
        try {
            closeable.close(false);   // true to force a close
        } catch (CloseVetoException e) {
            System.out.println("Close was vetoed "+e);
        }
    } 

    public static XComponent addonInitialize(XComponentContext addonXcc) {
        xcc = addonXcc;
        if (xcc == null) {
            System.out.println("Could not access component context");
            return null;
        }
        mcFactory = xcc.getServiceManager();
        if (mcFactory == null) {
            System.out.println("Office Service Manager is unavailable");
            return null;
        }
        try {
            Object oDesktop = mcFactory.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", xcc);
            xDesktop = Lo.qi(XDesktop.class, oDesktop);
        } catch (Exception e) {
            System.out.println("Could not access desktop");
            return null;
        }
        XComponent doc = xDesktop.getCurrentComponent();
        if (doc == null) {
            System.out.println("Could not access document");
            return null;
        }
        msFactory = Lo.qi(XMultiServiceFactory.class, doc);
        return doc;
    } 

    public static XComponent scriptInitialize(XScriptContext sc) {
        if (sc == null) {
            System.out.println("Script Context is null");
            return null;
        }
        xcc = sc.getComponentContext();
        if (xcc == null) {
            System.out.println("Could not access component context");
            return null;
        }
        mcFactory = xcc.getServiceManager();
        if (mcFactory == null) {
            System.out.println("Office Service Manager is unavailable");
            return null;
        }
        xDesktop = sc.getDesktop();
        if (xDesktop == null) {
            System.out.println("Could not access desktop");
            return null;
        }
        XComponent doc = xDesktop.getCurrentComponent();
        if (doc == null) {
            System.out.println("Could not access document");
            return null;
        }
        msFactory = Lo.qi(XMultiServiceFactory.class, doc);
        return doc;
    }  

    public static boolean dispatchCmd(String cmd) {
        return dispatchCmd(xDesktop.getCurrentFrame(), cmd, null);
    }

    public static boolean dispatchCmd(String cmd, PropertyValue[] props) {
        return dispatchCmd(xDesktop.getCurrentFrame(), cmd, props);
    }

    public static boolean dispatchCmd(XFrame frame, String cmd, PropertyValue[] props) {
        XDispatchHelper helper = createInstanceMCF(
                        XDispatchHelper.class, 
                        "com.sun.star.frame.DispatchHelper");
        if (helper == null) {
            System.out.println("Could not create dispatch helper for command " + cmd);
            return false;
        }
        try {
            XDispatchProvider provider = Lo.qi(XDispatchProvider.class, frame);
            helper.executeDispatch(provider, (".uno:" + cmd), "", 0, props);
            return true;
        } catch (java.lang.Exception e) {
            System.out.println("Could not dispatch \"" + cmd + "\":\n  " + e);
        }
        return false;
    } 

    public static String makeUnoCmd(String itemName) {
        return "vnd.sun.star.script:Foo/Foo." + itemName
                + "?language=Java&location=share";
    }

    public static String extractItemName(String unoCmd) {
        int fooPos = unoCmd.indexOf("Foo.");
        if (fooPos == -1) {
            System.out.println("Could not find Foo header in command: \"" + unoCmd + "\"");
            return null;
        }
        int langPos = unoCmd.indexOf("?language");
        if (langPos == -1) {
            System.out.println("Could not find language header in command: \"" + unoCmd + "\"");
            return null;
        }
        return unoCmd.substring(fooPos + 4, langPos);
    }  

    public static void inspect(Object obj) {
        if ((xcc == null) || (mcFactory == null)) {
            System.out.println("No office connection found");
            return;
        }
        try {
            Type[] ts = Info.getInterfaceTypes(obj);   // get class name for title
            String title = "Object";
            if ((ts != null) && (ts.length > 0)) title = ts[0].getTypeName() + " " + title;
            Object inspector = mcFactory.createInstanceWithContext(
                    "org.openoffice.InstanceInspector", xcc);
            if (inspector == null) {
                System.out.println("Inspector Service could not be instantiated");
                return;
            }
            System.out.println("Inspector Service instantiated");
            XIntrospection intro = createInstanceMCF(XIntrospection.class,
                    "com.sun.star.beans.Introspection");
            XIntrospectionAccess introAcc = intro.inspect(inspector);
            XIdlMethod method = introAcc.getMethod("inspect", -1);   // get ref to XInspector.inspect()
            System.out.println("inspect() method was found: " + (method != null));
            Object[][] params = new Object[][]{new Object[]{obj, title}};
            method.invoke(inspector, params);
        } catch (Exception e) {
            System.out.println("Could not access Inspector: " + e);
        }
    } 

    public static void mriInspect(Object obj) {
        XIntrospection xi = createInstanceMCF(XIntrospection.class, "mytools.Mri");
        if (xi == null) {
            System.out.println("MRI Inspector Service could not be instantiated");
            return;
        }
        System.out.println("MRI Inspector Service instantiated");
        xi.inspect(obj);
    }  

    public static int getColorInt(java.awt.Color color) {
        if (color == null) {
            System.out.println("No color supplied");
            return 0;
        } else return (color.getRGB() & 0xffffff);
    } 

    public static int hexString2ColorInt(String colStr) {
        java.awt.Color color = java.awt.Color.decode(colStr);
        return getColorInt(color);
    }

    public static String getColorHexString(java.awt.Color color) {
        if (color == null) {
            System.out.println("No color supplied");
            return "#000000";
        } else return int2HexString(color.getRGB() & 0xffffff);
    }  

    public static String int2HexString(int val) {
        String hex = Integer.toHexString(val);
        if (hex.length() < 6) hex = "000000".substring(0, 6 - hex.length()) + hex;
        return "#" + hex;
    }  

    public static void wait(int ms) {
        delay(ms);
    }

    public static void delay(int ms) {
        try {
            Thread.sleep(ms);
        } catch (InterruptedException e) {
            System.out.println(e);
        }
    }

    public static boolean isNullOrEmpty(String s) {
        return ((s == null) || (s.length() == 0));
    }

    public static void waitEnter() {
        System.out.println("Press Enter to continue...");
        try {
            System.in.read();
        } catch (java.io.IOException e) {
            System.out.println(e);
        }
    } 

    public static String getTimeStamp() {
        java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return sdf.format(new java.util.Date());
    }

    public static void printNames(String[] names) {
        printNames(names, 4);
    }

    public static void printNames(String[] names, int numPerLine) {
        if (names == null) System.out.println("  No names found");
        else {
            Arrays.sort(names, String.CASE_INSENSITIVE_ORDER);
            int nlCounter = 0;
            System.out.println("No. of names: " + names.length);
            for (String name : names) {
                System.out.print("  \"" + name + "\"");
                nlCounter++;
                if (nlCounter % numPerLine == 0) {
                    System.out.println();
                    nlCounter = 0;
                }
            }
            System.out.println("\n\n");
        }
    }  

    public static void printTable(String name, Object[][] table) {
        System.out.println("-- " + name + " ----------------");
        for (Object[] table1 : table) {
            for (Object table11 : table1) System.out.print("  " + table11);
            System.out.println();
        }
        System.out.println("-----------------------------\n");
    } 

    public static String capitalize(String s) {
        if ((s == null) || (s.length() == 0)) return null;
        else if (s.length() == 1) return s.toUpperCase();
        else return Character.toUpperCase(s.charAt(0)) + s.substring(1);
    } 

    public static int parseInt(String s) {
        if (s == null) return 0;
        try {
            return Integer.parseInt(s);
        } catch (NumberFormatException ex) {
            System.out.println(s + " could not be parsed as an int; using 0");
            return 0;
        }
    }  

    public static void addJar(String jarPath) {
        try {
            URLClassLoader classLoader
                    = (URLClassLoader) ClassLoader.getSystemClassLoader();
            java.lang.reflect.Method m
                    = URLClassLoader.class.getDeclaredMethod("addURL", java.net.URL.class);
            m.setAccessible(true);
            m.invoke(classLoader, new java.net.URL(jarPath));
        } catch (java.lang.NoSuchMethodException 
                | java.lang.SecurityException 
                | MalformedURLException 
                | java.lang.IllegalAccessException 
                | java.lang.IllegalArgumentException 
                | java.lang.reflect.InvocationTargetException e) {
            System.out.println(e);
        }
    }  

    public static String[] getContainerNames(XIndexAccess con) {
        if (con == null) {
            System.out.println("Container is null");
            return null;
        }
        int numElems = con.getCount();
        if (numElems == 0) {
            System.out.println("No elements in the container");
            return null;
        }
        ArrayList<String> namesList = new ArrayList<>();
        for (int i = 0; i < numElems; i++) {
            try {
                XNamed named = Lo.qi(XNamed.class, con.getByIndex(i));
                namesList.add(named.getName());
            } catch (Exception e) {
                System.out.println("Could not access name of element " + i);
            }
        }
        int sz = namesList.size();
        if (sz == 0) {
            System.out.println("No element names found in the container");
            return null;
        }
        String[] names = new String[sz];
        for (int i = 0; i < sz; i++) {
            names[i] = namesList.get(i);
        }
        return names;
    }  

    public static XPropertySet findContainerProps(XIndexAccess con, String nm) {
        if (con == null) {
            System.out.println("Container is null");
            return null;
        }
        for (int i = 0; i < con.getCount(); i++) {
            try {
                Object oElem = con.getByIndex(i);
                XNamed named = Lo.qi(XNamed.class, oElem);
                if (named.getName().equals(nm)) 
                    return (XPropertySet) Lo.qi(XPropertySet.class, oElem);
            } catch (Exception e) {
                System.out.println("Could not access element " + i + " " + e);
            }
        }
        System.out.println("Could not find a \"" + nm + "\" property set in the container");
        return null;
    }  

}

