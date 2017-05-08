package BlockEditMode;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import java.io.IOException;


public class EditingDocument {
    private static final String COM_SUN_STAR_FRAME_DESKTOP = "com.sun.star.frame.Desktop";
    private static final String NEW_DOCUMENT = "private:factory/swriter";

    private XComponentContext mxRemoteContext = null;
    private XMultiComponentFactory mxRemoteServiceManager = null;
    private Object desktop;
    private XTextDocument mxDoc;
    private XText mxDocText;

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        EditingDocument textDocuments1 = new EditingDocument();
        try {
            textDocuments1.runDemo();
        } catch (Exception e) {
            System.out.println(e.getMessage());
            e.printStackTrace();
        } finally {
            System.exit(0);
        }
    }

    private void runDemo() throws Exception {
        desktop = loadDesktopModule();
        java.io.File sourceFile = new java.io.File("PrintDemo.odt");
        StringBuilder sLoadFileUrl = new StringBuilder("file:///");
        sLoadFileUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
        editingExample();
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

    private XComponent loadFileFromURL(String url) throws IOException, com.sun.star.io.IOException {
        XComponentLoader xComponentLoader = UnoRuntime.queryInterface(XComponentLoader.class, desktop);
        PropertyValue[] loadProps = new PropertyValue[0];
        return xComponentLoader.loadComponentFromURL(
                url, "_blank", 0, loadProps);
    }

    /**
     * Sample for the various editing facilities described in the
     * developer's manual
     */
    private void editingExample() throws Exception {
        // create empty swriter document
        XComponent xEmptyWriterComponent = loadFileFromURL(NEW_DOCUMENT);
        // query its XTextDocument interface to get the text
        mxDoc = null;
        mxDoc = UnoRuntime.queryInterface(
                XTextDocument.class, xEmptyWriterComponent);

        // get a reference to the body text of the document
        mxDocText = null;
        mxDocText = mxDoc.getText();

        // Get a reference to the document's property set. This contains document
        // information like the current word count
        UnoRuntime.queryInterface(XPropertySet.class, mxDoc);

        // Simple text insertion example
        insertText();
    }

    // Setting the whole text of a document as one string
    private void insertText() {
        // Body Text and TextDocument example
        // demonstrate simple text insertion
        mxDocText.setString("This is the new body text of the document."
                + "\n\nThis is on the second line.\n\n");
    }
}


