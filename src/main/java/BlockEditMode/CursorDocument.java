package BlockEditMode;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XPageCursor;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import java.io.IOException;


public class CursorDocument {
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
        CursorDocument textDocuments1 = new CursorDocument();
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

        // Simple text insertion example
        insertText();
        viewCursorExample();
    }

    // Setting the whole text of a document as one string
    private void insertText() {
        // Body Text and TextDocument example
        // demonstrate simple text insertion
        mxDocText.setString("This is the new body text of the document."
                + "\n\nThis is on the second line.\n\n");
    }

    /** Sample for document changes, starting at the current view cursor position
     *  The sample changes the paragraph style and the character style at the
     *  current view cursor selection Open the sample file ViewCursorExampleFile,
     *  select some text and run the example.
     *  The current paragraph will be set to Quotations paragraph style.
     *  The selected text will be set to Quotation character style.
     */
    private void viewCursorExample() throws Exception {
        // query its XDesktop interface, we need the current component
        XDesktop xDesktop = UnoRuntime.queryInterface(
                XDesktop.class, desktop);
        // retrieve the current component and access the controller
        XComponent xCurrentComponent = xDesktop.getCurrentComponent();
        XModel xModel = UnoRuntime.queryInterface(XModel.class,
                xCurrentComponent);
        XController xController = xModel.getCurrentController();
        // the controller gives us the TextViewCursor
        XTextViewCursorSupplier xViewCursorSupplier =
                UnoRuntime.queryInterface(
                        XTextViewCursorSupplier.class, xController);
        XTextViewCursor xViewCursor = xViewCursorSupplier.getViewCursor();

        // query its XPropertySet interface, we want to set character and paragraph
        // properties
        XPropertySet xCursorPropertySet = UnoRuntime.queryInterface(
                XPropertySet.class, xViewCursor);
        // set the appropriate properties for character and paragraph style
        xCursorPropertySet.setPropertyValue("CharStyleName", "Quotation");
        xCursorPropertySet.setPropertyValue("ParaStyleName", "Quotations");
        // print the current page number
        XPageCursor xPageCursor = UnoRuntime.queryInterface(
                XPageCursor.class, xViewCursor);
        System.out.println("The current page number is " + xPageCursor.getPage());
        // the model cursor is much more powerful, so
        // we create a model cursor at the current view cursor position with the
        // following steps:
        // get the Text service from the TextViewCursor, it is an XTextRange:
        XText xDocumentText = xViewCursor.getText();
        // create a model cursor from the viewcursor
        XTextCursor xModelCursor = xDocumentText.createTextCursorByRange(
                xViewCursor.getStart());
        // now we could query XWordCursor, XSentenceCursor and XParagraphCursor
        // or XDocumentInsertable, XSortable or XContentEnumerationAccess
        // and work with the properties of com.sun.star.text.TextCursor
        // in this case we just go to the end of the paragraph and add some text.
        XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xModelCursor);
        // goto the end of the paragraph
        xParagraphCursor.gotoEndOfParagraph(false);
        xParagraphCursor.setString(" ***** Fin de semana! ******");
    }

}


