package BlockEditMode;

import com.sun.star.beans.IllegalTypeException;
import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyAttribute;
import com.sun.star.beans.PropertyExistException;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertyContainer;
import com.sun.star.beans.XPropertySet;
import com.sun.star.document.XDocumentPropertiesSupplier;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

import java.util.Collection;
import java.util.Collections;
import java.util.HashSet;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

public class DocumentUserProperties {

	private static final Logger LOGGER = Logger.getLogger(DocumentUserProperties.class.getName());
	private final XTextDocument textDocument;
	private Set<String> propertyNames;

	public DocumentUserProperties(XTextDocument document) {
		this.textDocument = document;
		
		propertyNames = new HashSet<String>();
		extractPropertyNames();
	}

	public void setString(String propertyName, String value) {
		setProperty(propertyName, value);
	}
	
	public String getString(String propertyName) {
		return (String) getProperty(propertyName);
	}
	
	public void setBoolean(String propertyName, Boolean value) {
		setProperty(propertyName, value);
	}
	
	public Boolean getBoolean(String propertyName) {
		return (Boolean) getProperty(propertyName);
	}
	
	private void setProperty(String propertyName, Object value) {
		XDocumentPropertiesSupplier propertiesSupplier = (XDocumentPropertiesSupplier) UnoRuntime.queryInterface(XDocumentPropertiesSupplier.class, textDocument);
		XPropertyContainer userDefinedProperties = propertiesSupplier.getDocumentProperties().getUserDefinedProperties();
		XPropertySet propertySet = UnoRuntime.queryInterface(XPropertySet.class, userDefinedProperties);
		try {
			userDefinedProperties.addProperty(propertyName, PropertyAttribute.REMOVEABLE, value);
		} catch (com.sun.star.lang.IllegalArgumentException ex) {
			LOGGER.log(Level.SEVERE, null, ex);
		} catch (PropertyExistException ex) {
			try {
				propertySet.setPropertyValue(propertyName, value);
			} catch (UnknownPropertyException ex1) {
				LOGGER.log(Level.SEVERE, null, ex1);
			} catch (PropertyVetoException ex1) {
				LOGGER.log(Level.SEVERE, null, ex1);
			} catch (com.sun.star.lang.IllegalArgumentException ex1) {
				LOGGER.log(Level.SEVERE, null, ex1);
			} catch (WrappedTargetException ex1) {
				LOGGER.log(Level.SEVERE, null, ex1);
			}
		} catch (IllegalTypeException ex) {
			LOGGER.log(Level.SEVERE, null, ex);
		}
	}

	private Object getProperty(String propertyName) {
		Object value = null;
		XDocumentPropertiesSupplier propertiesSupplier = (XDocumentPropertiesSupplier) UnoRuntime.queryInterface(XDocumentPropertiesSupplier.class, textDocument);
		XPropertyContainer userDefinedProperties = propertiesSupplier.getDocumentProperties().getUserDefinedProperties();
		XPropertySet propertySet = UnoRuntime.queryInterface(XPropertySet.class, userDefinedProperties);
		try {
			value = propertySet.getPropertyValue(propertyName);
		} catch (UnknownPropertyException ex) {
			//LOGGER.log(Level.SEVERE, null, ex);
		} catch (WrappedTargetException ex) {
			LOGGER.log(Level.SEVERE, null, ex);
		}
		return value;
	}
	
	public Collection<String> getPropertyNames() {
		return Collections.unmodifiableSet(propertyNames);
	}

	private void extractPropertyNames() {
		XDocumentPropertiesSupplier propertiesSupplier = (XDocumentPropertiesSupplier) UnoRuntime.queryInterface(XDocumentPropertiesSupplier.class, textDocument);
		XPropertyContainer userDefinedProperties = propertiesSupplier.getDocumentProperties().getUserDefinedProperties();
		XPropertySet propertySet = UnoRuntime.queryInterface(XPropertySet.class, userDefinedProperties);
		Property[] properties = propertySet.getPropertySetInfo().getProperties();
		if (properties != null) {
			for (Property property : properties) {
				propertyNames.add(property.Name);
			}
		}
	}
	
}
