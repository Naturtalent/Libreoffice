package it.naturtalent.libreoffice;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.accessibility.XAccessibleContext;
import com.sun.star.awt.XToolkit;
import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XIntrospection;
import com.sun.star.beans.XIntrospectionAccess;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.bridge.XBridge;
import com.sun.star.bridge.XBridgeFactory;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.connection.XConnection;
import com.sun.star.connection.XConnector;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XLayer;
import com.sun.star.drawing.XLayerManager;
import com.sun.star.drawing.XLayerSupplier;
import com.sun.star.drawing.framework.XConfiguration;
import com.sun.star.drawing.framework.XConfigurationController;
import com.sun.star.drawing.framework.XControllerManager;
import com.sun.star.drawing.framework.XModuleController;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XFramesSupplier;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.reflection.XIdlMethod;
import com.sun.star.ui.UIElementType;
import com.sun.star.ui.XSidebarPanel;
import com.sun.star.ui.XStatusbarItem;
import com.sun.star.ui.XUIConfigurationManager;
import com.sun.star.ui.XUIElement;
import com.sun.star.uno.Any;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;

import it.naturtalent.libreoffice.utils.GUI;
import it.naturtalent.libreoffice.utils.Info;
import it.naturtalent.libreoffice.utils.Lo;
import it.naturtalent.libreoffice.utils.Props;

public class Utils
{
	
	public enum OfficeTypes
	{
	  WRITER, CALC, DRAW, IMPRESS;

	  public String getType()
	  {
	    if ( this == WRITER)
	      return "swriter";
	    
	    if ( this == CALC)
		      return "scalc";
	    
	    if ( this == DRAW)
		      return "sdraw";

	    if ( this == IMPRESS)
		      return "simpress";

	    return null;
	  }
	}
	
	public enum LayerProperties
	{
	  Name, IsVisible, IsLocked;
	}
	
	

	
	// Remote office service Manager
	/*
	private static XMultiComponentFactory xServiceManager = null;
	public static XMultiServiceFactory getMultiServiceFactory()
	{
		if(xServiceManager == null)
		{			
			try
			{
				xContext = Bootstrap.bootstrap();
				xServiceManager = xContext.getServiceManager();
				
				XInterface xint= (XInterface)
						xServiceManager.createInstanceWithContext("com.sun.star.bridge.oleautomation.Factory",xContext);
				
				return (XMultiServiceFactory) UnoRuntime.queryInterface(
					      XMultiServiceFactory.class, xint);
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}			
		}
		
		return null;
	}
	*/
	
	public static XComponent newDocComponent(String docType) throws java.lang.Exception
	{
		String loadUrl = "private:factory/" + docType;
		
		XComponentContext xContext = getxContext();
		
		XMultiComponentFactory xMCF = xContext.getServiceManager();
		Object desktop = xMCF.createInstanceWithContext(
				"com.sun.star.frame.Desktop", xContext);
		
		XComponentLoader xComponentLoader = UnoRuntime.queryInterface(
				XComponentLoader.class, desktop);
		PropertyValue[] loadProps = new PropertyValue[0];
		return xComponentLoader.loadComponentFromURL(loadUrl, "_blank", 0,loadProps);
	}
	
   /** Load a document as template
    */
    public static XComponent newDocComponentFromTemplate(String loadUrl) throws java.lang.Exception
    {
    	
    	File sourceFile = new java.io.File(loadUrl);
    	StringBuffer sTemplateFileUrl = new StringBuffer("file:///");
        sTemplateFileUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
    	
        XComponentContext xContext = getxContext();
		XMultiComponentFactory xMCF = xContext.getServiceManager();
    	    	
        // retrieve the Desktop object, we need its XComponentLoader
        Object desktop = xMCF.createInstanceWithContext(
            "com.sun.star.frame.Desktop", xContext);
        
        XComponentLoader xComponentLoader = UnoRuntime.queryInterface(XComponentLoader.class, desktop);

        // define load properties according to com.sun.star.document.MediaDescriptor
        // the boolean property AsTemplate tells the office to create a new document
        // from the given file
        
        /*
        PropertyValue[] loadProps = new PropertyValue[1];
        loadProps[0] = new PropertyValue();
        loadProps[0].Name = "AsTemplate";
        loadProps[0].Value = new Boolean(true);
        */
        
        // load      
        PropertyValue[] loadProps = new PropertyValue[0];
        
        return xComponentLoader.loadComponentFromURL(sTemplateFileUrl.toString(), "_blank", 0, loadProps);
    }
    
	// local component context
	private static XComponentContext xContext = null;
	public static XComponentContext getxContext()
	{
		try
		{
			if(xContext == null)
				xContext = Bootstrap.bootstrap();
		} catch (BootstrapException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return xContext;
	}
    

    
    // Remote office service Manager
/*
private static XMultiComponentFactory xServiceManager = null;
public static XMultiServiceFactory getMultiServiceFactory()
{
	if(xServiceManager == null)
	{			
		try
		{
			xContext = Bootstrap.bootstrap();
			xServiceManager = xContext.getServiceManager();
			
			XInterface xint= (XInterface)
					xServiceManager.createInstanceWithContext("com.sun.star.bridge.oleautomation.Factory",xContext);
			
			return (XMultiServiceFactory) UnoRuntime.queryInterface(
				      XMultiServiceFactory.class, xint);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}			
	}
	
	return null;
}
*/

private static XComponentLoader officeComponentLoader = null;
public static XComponentLoader	getComponentLoader(String unoUrl) throws java.lang.Exception
{
	if (officeComponentLoader == null)
	{
		XComponentContext ctx = getxContext();

		// instantiate connector service
		Object x = ctx.getServiceManager().createInstanceWithContext(
				"com.sun.star.connection.Connector", ctx);

		XConnector xConnector = (XConnector) UnoRuntime.queryInterface(
				XConnector.class, x);

		// helper function to parse the UNO URL into a string array
		String a[] = parseUnoUrl(unoUrl);
		if (null == a)
		{
			throw new com.sun.star.uno.Exception("Couldn't parse UNO URL "
					+ unoUrl);
		}
		
		 // connect using the connection string part of the UNO URL only.
         XConnection connection = xConnector.connect(a[0]);

		// connect using the connection string part of the UNO URL only.
		x = ctx.getServiceManager().createInstanceWithContext(
				"com.sun.star.bridge.BridgeFactory", ctx);

		XBridgeFactory xBridgeFactory = (XBridgeFactory) UnoRuntime
				.queryInterface(XBridgeFactory.class, x);

		// create a nameless bridge with no instance provider
		// using the middle part of the UNO URL
		XBridge bridge = xBridgeFactory.createBridge("", a[1], connection,null);
		
		// query for the XComponent interface and add this as event listener
		XComponent xComponent = (XComponent) UnoRuntime.queryInterface(
				XComponent.class, bridge);
		//xComponent.addEventListener(this);

		// get the remote instance
		x = bridge.getInstance(a[2]);

		// Did the remote server export this object ?
		if (null == x)
		{
			throw new com.sun.star.uno.Exception(
					"Server didn't provide an instance for" + a[2], null);
		}

		// Query the initial object for its main factory interface
		XMultiComponentFactory xOfficeMultiComponentFactory = (XMultiComponentFactory) UnoRuntime
				.queryInterface(XMultiComponentFactory.class, x);

		// retrieve the component context (it's not yet exported from the
		// office)
		// Query for the XPropertySet interface.
		XPropertySet xProperySet = (XPropertySet) UnoRuntime
				.queryInterface(XPropertySet.class,
						xOfficeMultiComponentFactory);

		// Get the default context from the office server.
		Object oDefaultContext = xProperySet
				.getPropertyValue("DefaultContext");

		// Query for the interface XComponentContext.
		XComponentContext xOfficeComponentContext = (XComponentContext) UnoRuntime
				.queryInterface(XComponentContext.class, oDefaultContext);

		// now create the desktop service
		// NOTE: use the office component context here !
		Object oDesktop = xOfficeMultiComponentFactory
				.createInstanceWithContext("com.sun.star.frame.Desktop",
						xOfficeComponentContext);

		officeComponentLoader = (XComponentLoader) UnoRuntime
				.queryInterface(XComponentLoader.class, oDesktop);

		if (officeComponentLoader == null)
		{
			throw new com.sun.star.uno.Exception(
					"Couldn't instantiate com.sun.star.frame.Desktop", null);
		}
		
	}

	return officeComponentLoader;
}

	// Remote office service Manager
/*
private static XMultiComponentFactory xServiceManager = null;
public static XMultiServiceFactory getMultiServiceFactory()
{
	if(xServiceManager == null)
	{			
		try
		{
			xContext = Bootstrap.bootstrap();
			xServiceManager = xContext.getServiceManager();
			
			XInterface xint= (XInterface)
					xServiceManager.createInstanceWithContext("com.sun.star.bridge.oleautomation.Factory",xContext);
			
			return (XMultiServiceFactory) UnoRuntime.queryInterface(
				      XMultiServiceFactory.class, xint);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}			
	}
	
	return null;
}
*/

/** separates the uno-url into 3 different parts.
 */
protected static String[] parseUnoUrl(String url)
{
	String[] aRet = new String[3];

	if (!url.startsWith("uno:"))
	{
		return null;
	}

	int semicolon = url.indexOf(';');
	if (semicolon == -1)
		return null;

	aRet[0] = url.substring(4, semicolon);
	int nextSemicolon = url.indexOf(';', semicolon + 1);

	if (semicolon == -1)
		return null;
	aRet[1] = url.substring(semicolon + 1, nextSemicolon);

	aRet[2] = url.substring(nextSemicolon + 1);
	return aRet;
}

	/**
     * Zugriffexample: layer.getPropertyValue(LayerProperties.Name.name());
     * @param xDrawComponent
     * @return
     */
    public static List<XPropertySet> getLayerPropertySet(XComponent xDrawComponent)
	{
    	List<XPropertySet>propSet = new ArrayList<XPropertySet>();
    	
    	try
		{
			XLayerManager xLayerManager = null;
			XLayerSupplier xLayerSupplier = UnoRuntime.queryInterface(
			        XLayerSupplier.class, xDrawComponent);
			XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
			xLayerManager = UnoRuntime.queryInterface(XLayerManager.class, xNameAccess );
			
			XNameAccess nameAccess = (XNameAccess) UnoRuntime.queryInterface( 
					XNameAccess.class, xLayerManager);        
			String [] names = nameAccess.getElementNames();
			
			for(String name : names)
			{
				 Any any = (Any) nameAccess.getByName(name);
				 XPropertySet xLayerPropSet = UnoRuntime.queryInterface(XPropertySet.class, any);
				 propSet.add(xLayerPropSet);
			}
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    	
    	return propSet;
	}    
    

    public static Property[] propertiesInfo( XPropertySet xPropertySet)
	{
		 XPropertySetInfo info = xPropertySet.getPropertySetInfo();
		 return info.getProperties();
	}

    public static String [] printPropertyNames( XPropertySet xPropertySet)
	{
    	String [] names = null;
    	
		 Property [] props = propertiesInfo(xPropertySet);
		 for(Property property : props)
		 {
			 names = ArrayUtils.add(names, property.Name);
			 System.out.println(property.Name+" | "+property.Attributes);
		 }	
		 
		 return names;
	}

    public static void printPropertyValues( XPropertySet xPropertySet)
	{
    	try
		{
			String [] names = printPropertyNames(xPropertySet);
			if (ArrayUtils.isNotEmpty(names))
			{
				for (String name : names)
				{
					Object obj = xPropertySet.getPropertyValue(name);
					System.out.println(name+"  :  "+obj.getClass().getName()+"  :  "+obj);
				}
			}
		} catch (UnknownPropertyException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
    
    public static void testLayerProperty(XComponent xComponent) throws UnknownPropertyException, WrappedTargetException
 	{
		XModel xM = UnoRuntime.queryInterface(XModel.class, xComponent);
		XController xC = xM.getCurrentController();
		XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xC);
		
		Any any  = (Any) xPropertySet.getPropertyValue("ActiveLayer");
		XLayer xLayer = UnoRuntime.queryInterface(XLayer.class, any);		
		XPropertySet xLayerPropSet = UnoRuntime.queryInterface(XPropertySet.class, xLayer);
		
		Utils.printPropertyValues(xLayerPropSet);
		
		String name = (String) xLayerPropSet.getPropertyValue("Name");
		
		System.out.println("Layer: "+name);

 	}
    
    public static void testDocumentSettings(XComponent xComponent)
	{
		try
		{
			XMultiServiceFactory xFactory = UnoRuntime.queryInterface(
					XMultiServiceFactory.class, xComponent);
			XInterface settings = (XInterface) xFactory
					.createInstance("com.sun.star.drawing.DocumentSettings");
			XPropertySet xPropertySet = UnoRuntime.queryInterface(
					XPropertySet.class, settings);
			
			Utils.printPropertyValues(xPropertySet);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
    
    public static void printConfiguration(XComponent xComponent)
	{
    	XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		
		XControllerManager controllerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XModuleController moduleController = controllerManager.getModuleController();
		XConfigurationController configurationController = controllerManager.getConfigurationController();
		//XResource resource = configurationController.getResource("private:resource/pane/CenterPane");
		XConfiguration configuration = configurationController.getCurrentConfiguration();
		
		 PropertyValue [] pvs = xModel.getArgs();
		 for(PropertyValue pv : pvs)
		 {
			 System.out.println(pv.Name+" "+pv.Value);
			 
			 if(StringUtils.equals(pv.Name, "DocumentService"))
			 {
				 /*
				 XPropertySet xFTPSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, pv.Value);
				 System.out.println(xFTPSet);
				// reference the control by the Name
			      XControl xFTControl = m_xDlgContainer.getControl(sName);
			      xFixedText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, xFTControl);
			      XWindow xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, xFTControl);
			      xWindow.addMouseListener(_xMouseListener);
			      */
			 }
		 }

	}
    
    /*
     * 
     * Services
     * 
     * 
     */
    
    
    public static boolean closeFrame(com.sun.star.frame.XFrame xFrame)
    {
        boolean bClosed = false;

        try
        {
            // first try the new way: use new interface XCloseable
            // It replace the deprecated XTask::close() and should be preferred ...
            // if it can be queried.
            com.sun.star.util.XCloseable xCloseable =
                UnoRuntime.queryInterface(
                com.sun.star.util.XCloseable.class, xFrame);
            if (xCloseable!=null)
            {
                // We deliver the ownership of this frame not to the (possible)
                // source which throw a CloseVetoException. We wish to have it
                // under our own control.
                try
                {
                    xCloseable.close(false);
                    bClosed = true;
                }
                catch( com.sun.star.util.CloseVetoException exVeto )
                {
                    bClosed = false;
                }
            }
            else
            {
                // OK: the new way isn't possible. Try the old one.
                com.sun.star.frame.XTask xTask = UnoRuntime.queryInterface(com.sun.star.frame.XTask.class,
                                          xFrame);
                if (xTask!=null)
                {
                    // return value doesn't interest here. Because
                    // we forget this task ...
                    bClosed = xTask.close();
                }
            }
        }
        catch (com.sun.star.lang.DisposedException exDisposed)
        {
            // Of course - this task can be already dead - means disposed.
            // But for us it's not important. Because we tried to close it too.
            // And "already disposed" or "closed" should be the same ...
            bClosed = true;
        }

        return bClosed;
    }

    
    /**
     * OfficeDocument Type anzeigen
     * @param xComponent
     */
    public static void checkDocumentType(XComponent xComponent)
	{
    	XServiceInfo serviceInfo = (XServiceInfo)UnoRuntime.queryInterface(XServiceInfo.class, xComponent);
    	if(serviceInfo.supportsService("com.sun.star.drawing.GenericDrawingDocument"))
    		System.out.println("com.sun.star.drawing.GenericDrawingDocument");
    	else
    	{
    		if(serviceInfo.supportsService("com.sun.star.text.GenericTextDocument"))
    			System.out.println("com.sun.star.text.GenericTextDocument");
    		else
    		{
        		if(serviceInfo.supportsService("com.sun.star.sheet.SpreadsheetDocument"))
        			System.out.println("com.sun.star.sheet.SpreadsheetDocument");
        		else
        		{
            		if(serviceInfo.supportsService("com.sun.star.presentation.PresentationDocument"))
            			System.out.println("com.sun.star.presentation.PresentationDocument");
            		else
            		{
            			System.out.println("unklarer Dokumenttyp");
            		}
        		}    			
    		}    	
    	}
	}
    

	
	public static void printMSFServiceNames(XMultiServiceFactory msFactory )
	{
		String [] serviceNames = msFactory.getAvailableServiceNames();
		for(String name : serviceNames)
			System.out.println(name);
	}
	
	public static void printMCFServiceNames(XMultiComponentFactory mcFactory )
	{
		String [] serviceNames = mcFactory.getAvailableServiceNames();
		for(String name : serviceNames)
			System.out.println(name);
	}
    

	/**
	 * Alle Servicenames anzeigen
	 * @param xComponent
	 */
	public static void printAllServiceNames(XComponentContext xContext)
	{
		String [] serviceNames = xContext.getServiceManager().getAvailableServiceNames();
		for(String serviceName : serviceNames)
			System.out.println(serviceName);		
	}

    
	/**
	 * Alle Servicenames anzeigen
	 * @param xComponent
	 */
	public static void printSupportedServiceNames(XComponent xComponent)
	{
		String [] serviceNames = Utils.getSupportedServiceNames(xComponent);
		for(String serviceName : serviceNames)
			System.out.println(serviceName);
		
	}

	/**
	 * Namen der unterstuetzen Services
	 * 
	 * @param xComponent
	 * @return
	 */
	public static String [] getSupportedServiceNames(XComponent xComponent) 
	{
		String [] supportedServiceNames = new String [] {};
		
		XServiceInfo serviceInfo = (XServiceInfo)UnoRuntime.queryInterface(XServiceInfo.class, xComponent);
		if(serviceInfo != null)
			supportedServiceNames = serviceInfo.getSupportedServiceNames(); 
		
		return supportedServiceNames;
	}
	
	/**
	 * Service Properties anzeigen
	 * 
	 * @param xComponent
	 */
	public static void printServiceProperties(XComponentContext xContext, String serviceName)
	{
		System.out.println("\n\n Service: "+serviceName+" Properties\n");		
		try
		{
			Object serviceObj = xContext.getServiceManager().createInstanceWithContext(serviceName, xContext);
			if(serviceObj != null)
			{
				XPropertySet servicePropertySet = UnoRuntime.queryInterface(XPropertySet.class, serviceObj);
				if(servicePropertySet != null)
					Utils.printPropertyValues(servicePropertySet);
				else
					System.out.println("keine Properties");
			}
		} catch (com.sun.star.uno.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
	}

	
	/**
	 * Service OfficeDocument Properties anzeigen
	 * 
	 * @param xComponent
	 */
	public static void printOfficeDocumentProperties(XComponentContext xContext)
	{
		System.out.println("\n\n Service OfficeDocument Properties\n");
		Object officeDocument;
		try
		{
			officeDocument = xContext.getServiceManager().createInstanceWithContext(
					"com.sun.star.document.OfficeDocument", xContext);
			if(officeDocument != null)
			{
				XPropertySet officeDocumentPropSet = UnoRuntime
						.queryInterface(XPropertySet.class, officeDocument);
				Utils.printPropertyValues(officeDocumentPropSet);
			}
		} catch (com.sun.star.uno.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
	}
	
	public static void printUIElements(XFrame xFrame)
	{
		XPropertySet propSet = Lo.qi(XPropertySet.class, xFrame);
		XLayoutManager lm;
		try
		{
			lm = UnoRuntime.queryInterface(XLayoutManager.class, propSet.getPropertyValue("LayoutManager"));
			XUIElement[] uiElems = lm.getElements();
		      System.out.println("No. of UI Elements: " + uiElems.length);
		      for(XUIElement uiElem : uiElems)
		      {
		        System.out.println("  " + uiElem.getResourceURL() + "; " +
		                                  GUI.getUIElementTypeStr(uiElem.getType()) + "  URL: "+uiElem.getResourceURL());
		      }
		} catch (UnknownPropertyException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
	}
	
	public static XUIElement getXUIElement(XFrame xFrame, int toolType)
	{
		XPropertySet propSet = Lo.qi(XPropertySet.class, xFrame);
		XLayoutManager lm;
		try
		{
			lm = UnoRuntime.queryInterface(XLayoutManager.class, propSet.getPropertyValue("LayoutManager"));
			XUIElement[] uiElems = lm.getElements();
		      System.out.println("No. of UI Elements: " + uiElems.length);
		      for(XUIElement uiElem : uiElems)
		      {
		    	if(uiElem.getType() == toolType)
		    			return uiElem;
		      }
		} catch (UnknownPropertyException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		
		return null;
	}
	
	
	/*
	 * 
	 * 
	 * 
	 */
	
	public static XToolkit getXToolkit(XComponentContext xContext)
	{
		XToolkit xToolkit = null;
		
		try
		{
			xToolkit = UnoRuntime.queryInterface(
			        com.sun.star.awt.XToolkit.class,
			        xContext.getServiceManager().createInstanceWithContext(
			            "com.sun.star.awt.Toolkit", xContext));
		} catch (com.sun.star.uno.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return xToolkit;		
	}
	
	public static XFramesSupplier getXFramesSupplier(XComponentContext xContext)
	{
		XFramesSupplier xSupplier = null;
		
		try
		{
			xSupplier = UnoRuntime.queryInterface(
					com.sun.star.frame.XFramesSupplier.class,
					xContext.getServiceManager().createInstanceWithContext(
							"com.sun.star.frame.Desktop", xContext));
		} catch (com.sun.star.uno.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return xSupplier;		
	}

	public static XIndexAccess getXIndexAccess(XFramesSupplier xSupplier)
	{
		XIndexAccess xContainer = null;

		xContainer = UnoRuntime.queryInterface(
				com.sun.star.container.XIndexAccess.class,
				xSupplier.getFrames());

		return xContainer;
	}
	
	/**
	 * Properties der XDrawPage zeigen
	 * 
	 * @param xComponent
	 */
	public static void showXDrawPageProperty(XComponent xComponent)
	{
		 try
		{
			XDrawPagesSupplier oDPS = (XDrawPagesSupplier)
			            UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent);
			 XDrawPages the_pages = oDPS.getDrawPages();
			 XIndexAccess oDPi = (XIndexAccess)
			            UnoRuntime.queryInterface(XIndexAccess.class,the_pages);
			XDrawPage xDrawPage = (XDrawPage) AnyConverter.toObject(
			     new Type(XDrawPage.class),oDPi.getByIndex(0));
			
			XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, xDrawPage);
			Props.showProps("Page", props);
		} catch (IllegalArgumentException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IndexOutOfBoundsException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void showStatusbarUIElementsProperties(XFrame xFrame)
	{
		XUIElement uiElement = Utils.getXUIElement(xFrame, UIElementType.STATUSBAR);						
		XPropertySet propSet = Lo.qi(XPropertySet.class, uiElement);
		Props.showProps("UIElement", propSet);			
	}
	
	public static void showStatusbarConfiguration(XFrame xFrame, XComponentContext xContext)
	{
		try
		{
			XUIElement uiElement = Utils.getXUIElement(xFrame, UIElementType.STATUSBAR);	
			
			XInterface obj = (XInterface) uiElement.getRealInterface();
			XPropertySet propSet = Lo.qi(XPropertySet.class, uiElement);
			XUIConfigurationManager configManager = UnoRuntime.queryInterface(
					XUIConfigurationManager.class,
					propSet.getPropertyValue("ConfigurationSource"));					
			XIndexAccess xIndexAccess = configManager.getSettings("private:resource/statusbar/statusbar", true);			
			int n = xIndexAccess.getCount();
			
			// Type der Elemente: 'Type[[]com.sun.star.beans.PropertyValue]'
			//System.out.println("Elementtyp: "+xIndexAccess.getElementType());
			
			//PropertyValue pv = ((PropertyValue[]) xIndexAccess.getByIndex(0))[0];
			//XInterface o = Lo.qi(XInterface.class, pv);
			
			
			for(int i = 0;i < n;i++)			
				Props.showProps("Statusbar: "+i, (PropertyValue[]) xIndexAccess.getByIndex(i));
		} catch (IllegalArgumentException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (UnknownPropertyException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (NoSuchElementException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IndexOutOfBoundsException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void inspect(XComponentContext xContext, Object obj)
	/*
	 * call XInspector.inspect() in the Inspector.oxt extension Available from
	 * https://wiki.openoffice.org/wiki/Object_Inspector
	 */
	{
		if ((xContext == null))
		{
			System.out.println("No office connection found");
			return;
		}

		try
		{
			XMultiComponentFactory xMCF = xContext.getServiceManager();

			Type[] ts = Info.getInterfaceTypes(obj); // get class name for title
			String title = "Object";
			if ((ts != null) && (ts.length > 0))
				title = ts[0].getTypeName() + " " + title;

			Object inspector = xMCF.createInstanceWithContext(
					"org.openoffice.InstanceInspector", xContext);
			// hangs on second use
			if (inspector == null)
			{
				System.out
						.println("Inspector Service could not be instantiated");
				return;
			}

			System.out.println("Inspector Service instantiated");
			/*
			 * // report on inspector XServiceInfo si =
			 * Lo.qi(XServiceInfo.class, inspector);
			 * System.out.println("Implementation name: " +
			 * si.getImplementationName()); String[] serviceNames =
			 * si.getSupportedServiceNames(); for(String nm : serviceNames)
			 * System.out.println("Service name: " + nm);
			 */
			XIntrospection intro = DrawDocumentUtils.createInstanceMCF(xContext,
					XIntrospection.class, "com.sun.star.beans.Introspection");
			XIntrospectionAccess introAcc = intro.inspect(inspector);
			XIdlMethod method = introAcc.getMethod("inspect", -1); // get ref to
																	// XInspector.inspect()
			/*
			 * // alternative, low-level way of getting the method Object
			 * coreReflect = mcFactory.createInstanceWithContext(
			 * "com.sun.star.reflection.CoreReflection", xcc); XIdlReflection
			 * idlReflect = Lo.qi(XIdlReflection.class, coreReflect); XIdlClass
			 * idlClass =
			 * idlReflect.forName("org.openoffice.XInstanceInspector");
			 * XIdlMethod[] methods = idlClass.getMethods();
			 * System.out.println("No of methods: " + methods.length);
			 * for(XIdlMethod m : methods) System.out.println("  " +
			 * m.getName());
			 * 
			 * XIdlMethod method = idlClass.getMethod("inspect");
			 */
			System.out
					.println("inspect() method was found: " + (method != null));

			Object[][] params = new Object[][]
				{ new Object[]
							{ obj, title } };
			method.invoke(inspector, params);
		} catch (Exception e)
		{
			System.out.println("Could not access Inspector: " + e);
		}
	} // end of accessInspector()

	public static void showAccessibleContext(XAccessibleContext accessibleContext)
	{
		int n = accessibleContext.getAccessibleChildCount();
		for(int i = 0;i < n;i++)
		{
			try
			{
				XAccessible accessibleChild = accessibleContext.getAccessibleChild(i);
				XAccessibleContext xACChildChild = accessibleChild.getAccessibleContext();
				String name = xACChildChild.getAccessibleName();
				String desc = xACChildChild.getAccessibleDescription();
				System.out.println(name+" | "+desc);
			} catch (IndexOutOfBoundsException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
}
