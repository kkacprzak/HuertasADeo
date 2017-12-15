package fivedots;

/*  GUI.java
    Andrew Davison, ad@fivedots.coe.psu.ac.th, August 2016
    Heavily edited and formated by K. Kacprzak 2017
    A growing collection of utility functions to make Office easier to use. 
    They are currently divided into the following groups:
    * toolbar addition
    * floating frame, message box
    * controller and frame
    * Office container window
    * zooming
    * UI config manager
    * layout manager
    * menu bar
 */
import com.sun.star.accessibility.XAccessible;
import com.sun.star.accessibility.XAccessibleContext;
import com.sun.star.awt.PosSize;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.VclWindowPeerAttribute;
import com.sun.star.awt.WindowAttribute;
import com.sun.star.awt.WindowClass;
import com.sun.star.awt.WindowDescriptor;
import com.sun.star.awt.XExtendedToolkit;
import com.sun.star.awt.XMenuBar;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XSystemDependentWindowPeer;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XUserInputInterception;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.ElementExistException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchProviderInterception;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XFrames;
import com.sun.star.frame.XFramesSupplier;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XModel;
import com.sun.star.graphic.XGraphic;
import com.sun.star.lang.IllegalAccessException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.SystemDependent;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.ui.UIElementType;
import com.sun.star.ui.XImageManager;
import com.sun.star.ui.XModuleUIConfigurationManagerSupplier;
import com.sun.star.ui.XUIConfigurationManager;
import com.sun.star.ui.XUIConfigurationManagerSupplier;
import com.sun.star.ui.XUIElement;
import com.sun.star.view.DocumentZoomType;
import com.sun.star.view.XControlAccess;
import com.sun.star.view.XSelectionSupplier;
import java.util.ArrayList;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPasswordField;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

public class GUI {

    public static final short OPTIMAL = DocumentZoomType.OPTIMAL;
    public static final short PAGE_WIDTH = DocumentZoomType.PAGE_WIDTH;
    public static final short ENTIRE_PAGE = DocumentZoomType.ENTIRE_PAGE;
    public static final String MENU_BAR = "private:resource/menubar/menubar";
    public static final String STATUS_BAR = "private:resource/statusbar/statusbar";
    public static final String FIND_BAR = "private:resource/toolbar/findbar";
    public static final String STANDARD_BAR = "private:resource/toolbar/standardbar";
    public static final String TOOL_BAR = "private:resource/toolbar/toolbar";
    private static final String[] TOOBAR_NMS = {
        "3dobjectsbar", "addon_LibreLogo.OfficeToolBar",
        "alignmentbar", "arrowsbar", "arrowshapes", "basicshapes",
        "bezierobjectbar", "calloutshapes", "changes", "choosemodebar",
        "colorbar", "commentsbar", "commontaskbar", "connectorsbar",
        "custom_toolbar_1", "datastreams", "designobjectbar", "dialogbar",
        "drawbar", "drawingobjectbar", "drawobjectbar", "drawtextobjectbar",
        "ellipsesbar", "extrusionobjectbar", "findbar", "flowchartshapes",
        "fontworkobjectbar", "fontworkshapetype", "formatobjectbar",
        "Formatting", "formcontrols", "formcontrolsbar", "formdesign",
        "formobjectbar", "formsfilterbar", "formsnavigationbar",
        "formtextobjectbar", "frameobjectbar", "fullscreenbar",
        "gluepointsobjectbar", "graffilterbar", "graphicobjectbar",
        "insertbar", "insertcellsbar", "insertcontrolsbar",
        "insertobjectbar", "linesbar", "macrobar", "masterviewtoolbar",
        "mediaobjectbar", "moreformcontrols", "navigationobjectbar",
        "numobjectbar", "oleobjectbar", "optimizetablebar", "optionsbar",
        "outlinetoolbar", "positionbar", "previewbar", "previewobjectbar",
        "queryobjectbar", "rectanglesbar", "reportcontrols",
        "reportobjectbar", "resizebar", "sectionalignmentbar",
        "sectionshrinkbar", "slideviewobjectbar", "slideviewtoolbar",
        "sqlobjectbar", "standardbar", "starshapes", "symbolshapes",
        "tableobjectbar", "textbar", "textobjectbar", "toolbar",
        "translationbar", "viewerbar", "zoombar",};

    public static String getToobarResource(String nm) {
        for (String resNm : TOOBAR_NMS) {
            if (resNm.contains(nm)) {
                String resource = "private:resource/toolbar/" + resNm;
                System.out.println("Matched " + nm + " to " + resource);
                return resource;
            }
        }
        System.out.println("No matching resource for " + nm);
        return null;
    }  

    public static void addItemToToolbar(
            XComponent doc,
            String toolbarName,
            String itemName, String imFnm) 
            throws ElementExistException, 
            IllegalArgumentException, 
            IllegalAccessException,
            com.sun.star.container.NoSuchElementException,
            IndexOutOfBoundsException,
            WrappedTargetException {
        String cmd = Lo.makeUnoCmd(itemName);
        XUIConfigurationManager confMan = GUI.getUIConfigManagerDoc(doc);
        if (confMan == null) {
            System.out.println("Cannot create configuration manager");
            return;
        }
        XImageManager imageMan = Lo.qi(
                XImageManager.class,
                confMan.getImageManager());
        String[] cmds = {cmd};
        XGraphic[] pics = new XGraphic[1];
        pics[0] = Images.loadGraphicFile(imFnm);
        imageMan.insertImages((short) 0, cmds, pics);
        XIndexAccess settings = confMan.getSettings(toolbarName, true);
        XIndexContainer conSettings = Lo.qi(XIndexContainer.class, settings);
        PropertyValue[] itemProps = Props.makeBarItem(cmd, itemName);
        conSettings.insertByIndex(0, itemProps);
        confMan.replaceSettings(toolbarName, conSettings);
    } 

    public static XFrame createFloatingFrame(String title, int x, int y, int width, int height) 
            throws IllegalArgumentException {
        XToolkit xToolkit = Lo.createInstanceMCF(
                XToolkit.class,
                "com.sun.star.awt.Toolkit");
        if (xToolkit == null) return null;
        WindowDescriptor desc = new WindowDescriptor();
        desc.Type = WindowClass.TOP;
        desc.WindowServiceName = "modelessdialog";
        desc.ParentIndex = -1;
        desc.Bounds = new Rectangle(x, y, width, height);
        desc.WindowAttributes
                = WindowAttribute.BORDER 
                + WindowAttribute.MOVEABLE
                + WindowAttribute.CLOSEABLE 
                + WindowAttribute.SIZEABLE
                + VclWindowPeerAttribute.CLIPCHILDREN;
        XWindowPeer xWindowPeer = xToolkit.createWindow(desc);
        XWindow window = Lo.qi(XWindow.class, xWindowPeer);
        XFrame xFrame = Lo.createInstanceMCF(
                XFrame.class, 
                "com.sun.star.frame.Frame");
        if (xFrame == null) {
            System.out.println("Could not create frame");
            return null;
        }
        xFrame.setName(title);
        xFrame.initialize(window);
        XFramesSupplier xFramesSup = Lo.qi(XFramesSupplier.class, Lo.getDesktop());
        XFrames xFrames = xFramesSup.getFrames();
        if (xFrames == null) System.out.println("Mo desktop frames found");
        else xFrames.append(xFrame);
        window.setVisible(true);
        return xFrame;
    }  

    public static void showMessageBox(String title, String message) throws IllegalArgumentException {
        XToolkit xToolkit = Lo.createInstanceMCF(
                XToolkit.class,
                "com.sun.star.awt.Toolkit");
        XWindow xWindow = getWindow();
        if ((xToolkit == null) || (xWindow == null)) return;
        XWindowPeer xPeer = Lo.qi(XWindowPeer.class, xWindow);
        WindowDescriptor desc = new WindowDescriptor();
        desc.Type = WindowClass.MODALTOP;
        desc.WindowServiceName = "infobox";
        desc.ParentIndex = -1;
        desc.Parent = xPeer;
        desc.Bounds = new Rectangle(0, 0, 300, 200);
        desc.WindowAttributes = 
                WindowAttribute.BORDER 
                | WindowAttribute.MOVEABLE
                | WindowAttribute.CLOSEABLE;
        XWindowPeer descPeer = xToolkit.createWindow(desc);
        if (descPeer != null) {
            XMessageBox msgBox = Lo.qi(XMessageBox.class, descPeer);
            if (msgBox != null) {
                msgBox.setCaptionText(title);
                msgBox.setMessageText(message);
                msgBox.execute();
            }
        }
    }  

    public static void showJMessageBox(String title, String message) {
        JOptionPane.showMessageDialog(
                null, message, title,
                JOptionPane.INFORMATION_MESSAGE);
    }  

    public static String getPassword(String title, String inputMsg) {
        JLabel jl = new JLabel(inputMsg);
        JPasswordField jpf = new JPasswordField(24);
        Object[] ob = {jl, jpf};
        int result = JOptionPane.showConfirmDialog(
                null, ob, title,
                JOptionPane.OK_CANCEL_OPTION);
        if (result == JOptionPane.OK_OPTION) return new String(jpf.getPassword());
        else return null;
    }  

    public static XController getCurrentController(Object odoc) {
        XComponent doc = Lo.qi(XComponent.class, odoc);
        XModel model = Lo.qi(XModel.class, doc);
        if (model == null) {
            System.out.println("Document has no data model");
            return null;
        }
        return model.getCurrentController();
    }

    public static XFrame getFrame(XComponent doc) {
        return getCurrentController(doc).getFrame();
    }

    public static XControlAccess getControlAccess(XComponent doc) {
        return Lo.qi(XControlAccess.class, getCurrentController(doc));
    }

    public static XUserInputInterception getUII(XComponent doc) {
        return Lo.qi(XUserInputInterception.class, getCurrentController(doc));
    }

    public static XSelectionSupplier getSelectionSupplier(Object odoc) {
        XComponent doc = Lo.qi(XComponent.class, odoc);
        return Lo.qi(XSelectionSupplier.class, getCurrentController(doc));
    }

    public static XDispatchProviderInterception getDPI(XComponent doc) {
        return Lo.qi(XDispatchProviderInterception.class, getFrame(doc));
    }

    public static XWindow getWindow() {
        XDesktop desktop = Lo.getDesktop();
        XFrame frame = desktop.getCurrentFrame();
        if (frame == null) {
            System.out.println("No current frame");
            return null;
        } else return frame.getContainerWindow();
    }  

    public static XWindow getWindow(XComponent doc) {
        return getCurrentController(doc).getFrame().getContainerWindow();
    }

    public static void setVisible(Object objDoc, boolean isVisible) {
        XComponent doc = Lo.qi(XComponent.class, objDoc);
        XWindow xWindow = getFrame(doc).getContainerWindow();
        xWindow.setVisible(isVisible);
        xWindow.setFocus();
    }  

    public static void setVisible(boolean isVisible) {
        XWindow xWindow = getWindow();
        if (xWindow != null) {
            xWindow.setVisible(isVisible);
            xWindow.setFocus();
        }
    }  

    public static void setSizeWindow(XComponent doc, int width, int height) {
        XWindow xWindow = getWindow(doc);
        Rectangle rect = xWindow.getPosSize();
        xWindow.setPosSize(rect.X, rect.Y, width, height - 30, (short) 15);
    } 

    public static void setPosSize(XComponent doc, int x, int y, int width, int height) {
        XWindow xWindow = getWindow(doc);
        xWindow.setPosSize(x, y, width, height, PosSize.POSSIZE);   
    } 

    public static Rectangle getPosSize(XComponent doc) {
        XWindow xWindow = getWindow(doc);
        return xWindow.getPosSize();
    }

    public static XTopWindow getTopWindow() {
        XExtendedToolkit tk = Lo.createInstanceMCF(
                XExtendedToolkit.class,
                "com.sun.star.awt.Toolkit");
        if (tk == null) {
            System.out.println("Toolkit not found");
            return null;
        }
        XTopWindow topWin = tk.getActiveTopWindow();
        if (topWin == null) {
            System.out.println("Could not find top window");
            return null;
        }
        return topWin;
    }  

    public static String getTitleBar() {
        XTopWindow topWin = getTopWindow();
        if (topWin == null) return null;
        XAccessible acc = Lo.qi(XAccessible.class, topWin);
        if (acc == null) {
            System.out.println("Top window not accessible");
            return null;
        }
        XAccessibleContext accContext = acc.getAccessibleContext();
        return accContext.getAccessibleName();
    } 

    public static Rectangle getScreenSize() {  
        java.awt.GraphicsEnvironment ge = 
                java.awt.GraphicsEnvironment.getLocalGraphicsEnvironment();
        java.awt.GraphicsDevice[] gDevs = ge.getScreenDevices();
        java.awt.GraphicsConfiguration gc = gDevs[0].getDefaultConfiguration();
        java.awt.Rectangle bounds = gc.getBounds();
        java.awt.Insets insets =
                java.awt.Toolkit.getDefaultToolkit().getScreenInsets(gc);
        Rectangle rect = new Rectangle();  
        rect.X = bounds.x + insets.left;
        rect.Y = bounds.y + insets.top;
        rect.Height = bounds.height - (insets.top + insets.bottom);
        rect.Width = bounds.width - (insets.left + insets.right);
        return rect;
    } 

    public static void printRect(Rectangle r) {
        System.out.println("Rectangle: (" + r.X + ", " + r.Y
                + "), " + r.Width + " -- " + r.Height);
    }

    public static int getWindowHandle(XComponent doc) {
        XWindow win = getWindow(doc);
        XSystemDependentWindowPeer winPeer = Lo.qi(XSystemDependentWindowPeer.class, win);
        int handle = (Integer) winPeer.getWindowHandle(
                new byte[8],
                SystemDependent.SYSTEM_WIN32);
        return handle;
    }  

    public static void setLookFeel() {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException 
                | InstantiationException 
                | java.lang.IllegalAccessException 
                | UnsupportedLookAndFeelException e) {
            System.out.println("Could not set look and feel "+e);
        }
    }  

    public static void zoom(XComponent doc, short view) {
        if (view == OPTIMAL) Lo.dispatchCmd("ZoomOptimal");
        else if (view == PAGE_WIDTH) Lo.dispatchCmd("ZoomPageWidth");
        if (view == ENTIRE_PAGE) Lo.dispatchCmd("ZoomPage");
        else {
            System.out.println("Did not recognize zoom view: " + view + "; using optimal");
            Lo.dispatchCmd("ZoomOptimal");
        }
        Lo.delay(500);
    }  

    public static void zoomValue(XComponent doc, int value) {
        String[] zoomLabels = {"Zoom.Value", "Zoom.ValueSet", "Zoom.Type"};
        Object[] zoomVals = {(short) value, 28703, 0};
        Lo.dispatchCmd("Zoom", Props.makeProps(zoomLabels, zoomVals));
        Lo.delay(500);
    }  

    public static XUIConfigurationManager getUIConfigManager(XComponent doc) {
        XModel xModel = Lo.qi(XModel.class, doc);
        XUIConfigurationManagerSupplier xSupplier
                = Lo.qi(XUIConfigurationManagerSupplier.class, xModel);
        return xSupplier.getUIConfigurationManager();
    }  

    public static XUIConfigurationManager getUIConfigManagerDoc(XComponent doc) {
        String docType = Info.docTypeString(doc);    
        XModuleUIConfigurationManagerSupplier xSupplier
                = Lo.createInstanceMCF(
                        XModuleUIConfigurationManagerSupplier.class,
                        "com.sun.star.ui.ModuleUIConfigurationManagerSupplier");
        XUIConfigurationManager configMan = null;
        try {
            configMan = xSupplier.getUIConfigurationManager(docType);
        } catch (Exception e) {
            System.out.println("Could not create a config manager using \"" + docType + "\"");
        }
        return configMan;
    }  

    public static void printUICmds(
            XUIConfigurationManager configMan,
            String uiElemName) {
        try {
            XIndexAccess settings = configMan.getSettings(uiElemName, true);
            int numSettings = settings.getCount();
            System.out.println("No. of elements in \"" + uiElemName + "\" toolbar: " + numSettings);
            for (int i = 0; i < numSettings; i++) {
                PropertyValue[] settingProps = Lo.qi(PropertyValue[].class, settings.getByIndex(i));
                Object val = Props.getValue("CommandURL", settingProps);
                System.out.println(i + ") " + Props.propValueToString(val));
            }
            System.out.println();
        } catch (com.sun.star.container.NoSuchElementException 
                | IllegalArgumentException 
                | IndexOutOfBoundsException 
                | WrappedTargetException e) {
            System.out.println(e);
        }
    }  

    public static void printUICmds(XComponent doc, String uiElemName) {
        XUIConfigurationManager configMan = GUI.getUIConfigManagerDoc(doc);
        if (configMan == null) 
            System.out.println("Cannot create configuration manager");
        else GUI.printUICmds(configMan, uiElemName);
    }  

    public static XLayoutManager getLayoutManager() {
        XDesktop desktop = Lo.getDesktop();
        XFrame frame = desktop.getCurrentFrame();
        if (frame == null) {
            System.out.println("No current frame");
            return null;
        }
        XLayoutManager lm = null;
        try {
            XPropertySet propSet = Lo.qi(XPropertySet.class, frame);
            lm = Lo.qi(XLayoutManager.class, propSet.getPropertyValue("LayoutManager"));
        } catch (UnknownPropertyException | WrappedTargetException e) {
            System.out.println("Could not access layout manager "+e);
        }
        return lm;
    } 

    public static XLayoutManager getLayoutManager(XComponent doc) {
        XLayoutManager lm = null;
        try {
            XPropertySet propSet = Lo.qi(XPropertySet.class, getFrame(doc));
            lm = Lo.qi(XLayoutManager.class, propSet.getPropertyValue("LayoutManager"));
        } catch (UnknownPropertyException | WrappedTargetException e) {
            System.out.println("Could not access layout manager "+e);
        }
        return lm;
    } 

    public static void printUIs() {
        printUIs(getLayoutManager());
    }

    public static void printUIs(XComponent doc) {
        printUIs(getLayoutManager(doc));
    }

    public static void printUIs(XLayoutManager lm) {
        if (lm == null) System.out.println("No layout manager found");
        else {
            XUIElement[] uiElems = lm.getElements();
            System.out.println("No. of UI Elements: " + uiElems.length);
            for (XUIElement uiElem : uiElems) {
                System.out.println("  " + uiElem.getResourceURL() + "; "
                        + getUIElementTypeStr(uiElem.getType()));
            }
            System.out.println();
        }
    }  

    public static String getUIElementTypeStr(short t) {
        if (t == UIElementType.UNKNOWN) return "unknown";
        if (t == UIElementType.MENUBAR) return "menubar";
        if (t == UIElementType.POPUPMENU) return "popup menu";
        if (t == UIElementType.TOOLBAR) return "toolbar";
        if (t == UIElementType.STATUSBAR) return "status bar";
        if (t == UIElementType.FLOATINGWINDOW) return "floating window";
        if (t == UIElementType.PROGRESSBAR) return "progress bar";
        if (t == UIElementType.TOOLPANEL) return "tool panel";
        if (t == UIElementType.DOCKINGWINDOW) return "docking window";
        if (t == UIElementType.COUNT) return "count";
        return "??";
    }  

    public static void printAllUICommands(XComponent doc) {
        XUIConfigurationManager confMan = getUIConfigManagerDoc(doc);
        if (confMan == null) {
            System.out.println("No configuration manager found");
            return;
        }
        XLayoutManager lm = getLayoutManager(doc);
        if (lm == null) {
            System.out.println("No layout manager found");
            return;
        }
        XUIElement[] uiElems = lm.getElements();
        System.out.println("No. of UI Elements: " + uiElems.length);
        String uiElemName;
        for (XUIElement uiElem : uiElems) {
            uiElemName = uiElem.getResourceURL();
            System.out.println("--- " + uiElemName + " ---");
            printUICmds(confMan, uiElemName);
        }
    }  

    public static void showOne(XComponent doc, String showElem) {
        ArrayList<String> showElems = new ArrayList<String>();
        showElems.add(showElem);
        showOnly(doc, showElems);
    }  

    public static void showOnly(XComponent doc, ArrayList<String> showElems) {
        XLayoutManager lm = getLayoutManager(doc);
        if (lm == null)  System.out.println("No layout manager found");
        else {
            XUIElement[] uiElems = lm.getElements();
            hideExcept(lm, uiElems, showElems);
            for (String elemName : showElems) {  // these elems are not in lm
                lm.createElement(elemName);       // so need to be created & shown
                lm.showElement(elemName);
                System.out.println(elemName + " made visible");
            }
        }
    }  

    public static void hideExcept(XLayoutManager lm, XUIElement[] uiElems,
            ArrayList<String> showElems) {
        for (XUIElement uiElem : uiElems) {
            String elemName = uiElem.getResourceURL();
            boolean toHide = true;
            for (int i = 0; i < showElems.size(); i++) {
                if (showElems.get(i).equals(elemName)) {
                    showElems.remove(i);  
                    toHide = false;       
                    break;
                }
            }
            if (toHide) {
                lm.hideElement(elemName);
                System.out.println(elemName + " hidden");
            }
        }
    }  

    static void showNone(XComponent doc) {
        XLayoutManager lm = getLayoutManager(doc);
        if (lm == null) System.out.println("No layout manager found");
        else {
            XUIElement[] uiElems = lm.getElements();
            for (XUIElement uiElem : uiElems) {
                String elemName = uiElem.getResourceURL();
                lm.hideElement(elemName);
                System.out.println(elemName + " hidden");
            }
        }
    } 

    public static XMenuBar getMenubar(XLayoutManager lm) {
        if (lm == null) {
            System.out.println("No layout manager available for menu discovery");
            return null;
        }
        XMenuBar bar = null;
        try {
            XUIElement oMenuBar = lm.getElement(GUI.MENU_BAR);
            XPropertySet props = Lo.qi(XPropertySet.class, oMenuBar);
            bar = Lo.qi(XMenuBar.class, props.getPropertyValue("XMenuBar"));
            if (bar == null) System.out.println("Menubar reference not found");
        } catch (UnknownPropertyException | WrappedTargetException e) {
            System.out.println("Could not access menubar "+e);
        }
        return bar;
    }  

    public static short getMenuMaxID(XMenuBar bar) {
        if (bar == null) return -1;
        short itemCount = bar.getItemCount();
        System.out.println("No items in menu bar: " + itemCount);
        short maxID = -1;
        for (short i = 0; i < itemCount; i++) {
            short id = bar.getItemId(i);
            if (id > maxID) maxID = id;
        }
        return maxID;
    }  

} 

