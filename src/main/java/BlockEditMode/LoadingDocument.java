package BlockEditMode;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import java.io.IOException;


public class LoadingDocument {
    private static final String COM_SUN_STAR_FRAME_DESKTOP = "com.sun.star.frame.Desktop";
    private static final String NEW_DOCUMENT = "private:factory/swriter";

    private XComponentContext mxRemoteContext = null;
    private XMultiComponentFactory mxRemoteServiceManager = null;

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        LoadingDocument textDocuments1 = new LoadingDocument();
        try {
            textDocuments1.runDemo();
        } catch (Exception e) {
            System.out.println(e.getMessage());
            e.printStackTrace();
        } finally {
            System.exit(0);
        }
    }

    protected void runDemo() throws Exception {
        Object desktopModule = loadDesktopModule();
        java.io.File sourceFile = new java.io.File("PrintDemo.odt");
        StringBuffer sLoadFileUrl = new StringBuffer("file:///");
        sLoadFileUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
        //  loadFileFromURL(desktopModule, sLoadFileUrl.toString());
        loadFileFromURL(desktopModule, sLoadFileUrl.toString());
    }

    private XMultiComponentFactory getRemoteServiceManager()
            throws Exception {
        if (mxRemoteContext == null && mxRemoteServiceManager == null) {
            // get the remote office context. If necessary a new office
            // process is started
            mxRemoteContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
            System.out.println("Connected to a running office ...");
            mxRemoteServiceManager = mxRemoteContext.getServiceManager();
        }
        return mxRemoteServiceManager;
    }

    private Object loadDesktopModule() throws Exception {
        // get the remote service manager
        mxRemoteServiceManager = this.getRemoteServiceManager();
        // retrieve the Desktop object, we need its XComponentLoader
        return mxRemoteServiceManager.createInstanceWithContext(
                COM_SUN_STAR_FRAME_DESKTOP, mxRemoteContext);
    }

    private XComponent loadFileFromURL(Object desktop, String url) throws IOException, com.sun.star.io.IOException {
        XComponentLoader xComponentLoader = UnoRuntime.queryInterface(XComponentLoader.class, desktop);
        PropertyValue[] loadProps = new PropertyValue[0];
        return xComponentLoader.loadComponentFromURL(
                url, "_blank", 0, loadProps);
    }

}


