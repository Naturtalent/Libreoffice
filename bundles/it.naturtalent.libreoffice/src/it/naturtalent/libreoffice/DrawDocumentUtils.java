package it.naturtalent.libreoffice;

import java.text.NumberFormat;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.swing.plaf.synth.SynthSpinnerUI;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import com.sun.star.accessibility.AccessibleRole;
import com.sun.star.accessibility.XAccessible;
import com.sun.star.accessibility.XAccessibleComponent;
import com.sun.star.accessibility.XAccessibleContext;
import com.sun.star.accessibility.XAccessibleText;
import com.sun.star.awt.Point;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.Size;
import com.sun.star.awt.XExtendedToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.SlideSorter;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XLayer;
import com.sun.star.drawing.XLayerManager;
import com.sun.star.drawing.XLayerSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.drawing.XSlideSorterBase;
import com.sun.star.drawing.framework.AnchorBindingMode;
import com.sun.star.drawing.framework.ConfigurationChangeEvent;
import com.sun.star.drawing.framework.ResourceId;
import com.sun.star.drawing.framework.XConfiguration;
import com.sun.star.drawing.framework.XConfigurationChangeListener;
import com.sun.star.drawing.framework.XConfigurationController;
import com.sun.star.drawing.framework.XControllerManager;
import com.sun.star.drawing.framework.XPane;
import com.sun.star.drawing.framework.XResource;
import com.sun.star.drawing.framework.XResourceId;
import com.sun.star.drawing.framework.XView;
import com.sun.star.frame.DispatchResultEvent;
import com.sun.star.frame.DispatchResultState;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;

import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.rendering.XCanvas;
import com.sun.star.rendering.XGraphicDevice;
import com.sun.star.uno.Any;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;
import com.sun.star.util.URL;
import com.sun.star.view.XSelectionSupplier;

import it.naturtalent.libreoffice.environment.example.OfficeConnect;
import it.naturtalent.libreoffice.listeners.MouseListener;
import it.naturtalent.libreoffice.utils.Lo;
import it.naturtalent.libreoffice.utils.Props;

public class DrawDocumentUtils
{
	// Systemlayernamen
	private static Map <String, String>localLayerNamesMap;
	
	// Layernamen werden intern in Locale Namen konvertiert (z.B. 'controls' in 'Steuerelemente'. 
	// Diese Tabelle unterstuetzt die Simulation dieser Konvertierung 
	private static String getLocalLayerName(String name)
	{
		if (localLayerNamesMap == null)
		{
			localLayerNamesMap = new HashMap<String, String>();
								
			localLayerNamesMap.put("layout", "Layout");
			localLayerNamesMap.put("background", "");
			localLayerNamesMap.put("backgroundobjects", "");
			localLayerNamesMap.put("controls", "Steuerelemente");
			localLayerNamesMap.put("measurelines", "Maßlinien");
		}
		
		return localLayerNamesMap.get(name);
	}

	// Systempagenamen
	private static Map <String, String>localPageNamesMap;
	
	// ersetzen in spaeteren Java-Versionen durch eine direkte Map-Initialisierung
	private static void initLocalPageNameMap()
	{
		if(localPageNamesMap == null)
		{
			localPageNamesMap = new HashMap<String, String>();
			localPageNamesMap.put("page", "Folie");
		}
	}
	

	/*
	 * 
	 * 
	 * 
	 */
	/*
	public static Double getScaleFactor(XComponent xComponent)
	{
		if (xComponent != null)
		{
			try
			{
				XMultiServiceFactory xFactory = UnoRuntime.queryInterface(
						XMultiServiceFactory.class, xComponent);
				XInterface settings = (XInterface) xFactory
						.createInstance("com.sun.star.drawing.DocumentSettings");
				XPropertySet xPageProperties = UnoRuntime.queryInterface(
						XPropertySet.class, settings);
				Integer scaleNumerator = (Integer) xPageProperties
						.getPropertyValue("ScaleNumerator");
				Integer scaleDenominator = (Integer) xPageProperties
						.getPropertyValue("ScaleDenominator");
				
				double scaleFactor = ((double)scaleDenominator/(double)scaleNumerator);
				return scaleFactor;

			} catch (UnknownPropertyException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WrappedTargetException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (com.sun.star.uno.Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return null;
	}
	*/

	
	
	/*
	 * Accessible Contexts
	 *  
	 * 
	 */

	/**
	 * Ueber XAccessibleComponent Zugriff auf Groesse und Position des TopWindows moeglich.
	 *  
	 * @param xContext
	 * @return
	 */
	public static XAccessibleComponent getAccessibleComponent(XComponentContext xContext)
	{
		XAccessibleComponent aComp = null;
		XTopWindow xTopWindow = getDrawDocumentXTopWindow(xContext);
		XAccessible xTopWindowAccessible = (XAccessible)
		         UnoRuntime.queryInterface(XAccessible.class, xTopWindow);	
		XAccessibleContext xAccessibleContext = xTopWindowAccessible.getAccessibleContext();
		aComp = (XAccessibleComponent) UnoRuntime.queryInterface(
              XAccessibleComponent.class, xAccessibleContext);	
		return aComp;
	}
	
	/**
	 * TopWindow des DrawDocumnts zurueckgeben
	 * 
	 * @param xContext
	 * @return
	 */
	public static XTopWindow getDrawDocumentXTopWindow(XComponentContext xContext)
	{
		XTopWindow [] xTopWindows = getXTopWindows(xContext);		
		for(XTopWindow xTopWindow : xTopWindows)
		{
			XAccessible xTopWindowAccessible = (XAccessible)
			         UnoRuntime.queryInterface(XAccessible.class, xTopWindow);				
			XAccessibleContext xAccessibleContext = xTopWindowAccessible.getAccessibleContext();
			String name = xAccessibleContext.getAccessibleName();
			if(StringUtils.contains(name, "Draw"))
				return xTopWindow;				
		}
		
		return null;
	}
	
	/**
	 * Alle im 'XAccessibleContext' gespeicherten TopWindows zurueckgeben.
	 * 
	 * @param xContext
	 * @return
	 */
	public static XTopWindow [] getXTopWindows(XComponentContext xContext)
	{
		XTopWindow [] xTopWindows;
		XExtendedToolkit tk = DrawDocumentUtils.createInstanceMCF(xContext, XExtendedToolkit.class, 
                "com.sun.star.awt.Toolkit");
		
		int n = tk.getTopWindowCount();
		xTopWindows = new XTopWindow [n]; 
		for(int i = 0;i < n;i++)
			try
			{
				xTopWindows[i]=tk.getTopWindow(i);
			} catch (IndexOutOfBoundsException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		
		return xTopWindows;
	}

	/**
	 * Den XAccessibleContext fuer das DrawDocument ermitteln.
	 * Sucht in allen XTopWindows nach dem XAccessibleContext dessen Name den String 'Draw' enthaelt.
	 *  
	 * @param xContext
	 * @return
	 */
	public static XAccessibleContext getAccessibleContext(XComponentContext xContext)
	{
		XExtendedToolkit tk = DrawDocumentUtils.createInstanceMCF(xContext, XExtendedToolkit.class, 
                "com.sun.star.awt.Toolkit");
		
		int n = tk.getTopWindowCount();
		for(int i = 0;i < n;i++)
		{
			try
			{
				XTopWindow topWindow = tk.getTopWindow(i);
				XAccessible xTopWindowAccessible = (XAccessible)
				         UnoRuntime.queryInterface(XAccessible.class, topWindow);				
				XAccessibleContext xAccessibleContext = xTopWindowAccessible.getAccessibleContext();
				String name = xAccessibleContext.getAccessibleName();
				if(StringUtils.contains(name, "Draw"))
				{					
					return xAccessibleContext;
				}
			} catch (IndexOutOfBoundsException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}				
		}
		
		return null;
	}
	
	/**
	 * Rueckgabe der im Statusbereich angezeigten MousePosition
	 * 
	 * @param parentAccessibleContext (DrawDocument)
	 * @return
	 */
	public static Double [] getStatusposition(XAccessibleContext parentAccessibleContext)
	{
		try
		{
			// Index '6' ist die Statusbar
			XAccessible statusBarAccessible = parentAccessibleContext.getAccessibleChild(6);
			XAccessibleContext statusBar = statusBarAccessible.getAccessibleContext();
			
			// Item[0] enthaelt die Cursorposition	
			XAccessible positionAccessible = statusBar.getAccessibleChild(1);
			XAccessibleContext position = positionAccessible.getAccessibleContext();
			
			int role = position.getAccessibleRole();
			if(role == AccessibleRole.LABEL)
			{
				 XInterface posInterface = (XInterface)
			                UnoRuntime.queryInterface(XInterface.class, position);
				 String [] pos = StringUtils.split(getString(posInterface));
				 
				Double[] point = new Double[2];
				try
				{
					NumberFormat nf = NumberFormat.getInstance(Locale.GERMAN);
					point[0] = nf.parse(pos[0]).doubleValue();
					point[1] = nf.parse(pos[2]).doubleValue();
				} catch (ParseException e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				} 				
				 return point;
			}
			
		} catch (IndexOutOfBoundsException e)
		{
			// TODO Auto-generated catch block
			//e.printStackTrace();
		}
		
		return null;
	}
	
	private static String getString(XInterface xInt)
	{
		XAccessibleText oText = (XAccessibleText) UnoRuntime
				.queryInterface(XAccessibleText.class, xInt);
		return oText.getText();
	}

	/*
	 * Utils
	 *  
	 * 
	 */

	public static <T> T createInstanceMSF(XComponent xComponent, Class<T> aType, String serviceName)
	{
		Object o;
		T interfaceObj = null;
		
		XMultiServiceFactory msFactory = UnoRuntime.queryInterface(XMultiServiceFactory.class, xComponent);
						
		try
		{
			o = msFactory.createInstance(serviceName);
			interfaceObj = Lo.qi(aType, o);    
		} catch (com.sun.star.uno.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}     		 
		
		return interfaceObj;
	}
	
	public static <T> T createInstanceMSF(Class<T> aType, String serviceName, XMultiServiceFactory msFactory)
	{
		if (msFactory == null)
		{
			System.out.println("No document found");
			return null;
		}
			
		T interfaceObj = null;
		try
		{
			// create service component
			Object o = msFactory.createInstance(serviceName); 
			interfaceObj = UnoRuntime.queryInterface(aType, o);
			
			// uses bridge to obtain proxy to remote interface inside service;
			// implements casting across process boundaries
		} catch (Exception e)
		{
			System.out.println("Couldn't create interface for \"" + serviceName+ "\":\n  " + e);
		}
		
		return interfaceObj;
	}

	public static <T> T createInstanceMCF(XComponentContext xContext, Class<T> aType, String serviceName)
	{
		 T interfaceObj = null;
		XMultiComponentFactory xMCF = xContext.getServiceManager();
		
		Object o;
		try
		{
			o = xMCF.createInstanceWithContext(serviceName, xContext);
			interfaceObj = UnoRuntime.queryInterface(aType, o);    
		} catch (com.sun.star.uno.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}     		 
		
		return interfaceObj;
	}
	
	/**
	 * @param sURL
	 * @return
	 */
	public static com.sun.star.util.URL parseURL(XComponentContext xContext, String sURL)
	{
		com.sun.star.util.URL aURL = null;

		if (sURL == null || sURL.equals(""))
		{
			System.out.println("wrong using of URL parser");
			return null;
		}

		try
		{
			// Create special service for parsing of given URL.
			com.sun.star.util.XURLTransformer xParser = UnoRuntime
					.queryInterface(com.sun.star.util.XURLTransformer.class,
							xContext.getServiceManager()
									.createInstanceWithContext(
											"com.sun.star.util.URLTransformer",
											xContext));

			// Because it's an in/out parameter we must use an array of URL
			// objects.
			com.sun.star.util.URL[] aParseURL = new com.sun.star.util.URL[1];
			aParseURL[0] = new com.sun.star.util.URL();
			aParseURL[0].Complete = sURL;

			// Parse the URL
			xParser.parseStrict(aParseURL);

			aURL = aParseURL[0];
		} catch (com.sun.star.uno.RuntimeException exRuntime)
		{
			// Any UNO method of this scope can throw this exception.
			// Reset the return value only.
		} catch (com.sun.star.uno.Exception exUno)
		{
			// "createInstance()" method of used service manager can throw it.
			// Then it wasn't possible to get the URL transformer.
			// Return default instead of really parsed URL.
		}

		return aURL;
	}
	
	/*
	 * 
	 * 
	 * 
	 * 
	 * 
	 */
	
	public static DispatchResultEvent dispatchCmd(XComponentContext xContext, XFrame frame, String cmd, PropertyValue[] props)
	// cmd does not include the ".uno:" substring; e.g. pass "Zoom" not ".uno:Zoom"
	{
		DispatchResultEvent dispatchResult = null;
		
	    XDispatchHelper helper = createInstanceMCF(xContext,XDispatchHelper.class, "com.sun.star.frame.DispatchHelper");
	    if (helper != null) 
		{
			try
			{
				XDispatchProvider provider = UnoRuntime.queryInterface(XDispatchProvider.class,frame);

				/*
				 * returns failure even when the event works (?), and an illegal
				 * value when the dispatch actually does fail
				 */
				Object res = helper.executeDispatch(provider, (".uno:" + cmd), "", 0,props);
				if (res instanceof DispatchResultEvent)
				{
					dispatchResult = (DispatchResultEvent) res;								
					if (dispatchResult.State == DispatchResultState.FAILURE)
						System.out.println("Dispatch failed for \"" + cmd + "\"");
					
					else if (dispatchResult.State == DispatchResultState.DONTKNOW)
						System.out.println("Dispatch result unknown for \"" + cmd + "\"");
				}

			} catch (java.lang.Exception e)
			{
				System.out
						.println("Could not dispatch \"" + cmd + "\":\n  " + e);
			}	
		}
	    
	    return dispatchResult;
	}  

	
	/*
	 * 
	 *  Resources
	 * 
	 */
	
	private static String drawingFrameworkPanes [] = {
			"private:resource/pane/CenterPane",
			"private:resource/pane/LeftImpressPane",
			"private:resource/pane/LeftDrawPane",
			"private:resource/pane/RightPane",
		};
	
	private static String drawingFrameworkViews [] = {
			"private:resource/view/ImpressView",
			"private:resource/view/GraphicView",
			"private:resource/view/OutlineView",
			"private:resource/view/NotesView",
			"private:resource/view/HandoutView",
			"private:resource/view/SlideSorter",
			"private:resource/view/TaskPane",
		};


	public static XConfiguration getConfiguration(XComponent xComponent)
	{
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XConfigurationController configurationController = xcontollerManager.getConfigurationController();
		return configurationController.getCurrentConfiguration();
	}
	
	public static XResource getResource(XComponent xComponent, XComponentContext xContext, String resourceURL)
	{
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XConfigurationController configurationController = xcontollerManager.getConfigurationController();
		XResourceId resourceID = ResourceId.create(xContext, resourceURL);
		XResource resource = configurationController.getResource(resourceID);
		return resource;
	}

	public static XPane getConiguredPane(XComponent xComponent,XComponentContext xContext, int index)
	{
		XPane pane = null;

		// Zugriff auf die Resourcen via XConfigurationController
		if (index < drawingFrameworkPanes.length)
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
			XController xController = xModel.getCurrentController();
			XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
			XConfigurationController configurationController = xcontollerManager.getConfigurationController();
			XConfiguration configuration = configurationController.getCurrentConfiguration();
			XResourceId resourceIDempty = ResourceId.createEmpty(xContext);
			XResourceId[] resourceIDs = configuration.getResources(
					resourceIDempty, drawingFrameworkPanes[index],
					AnchorBindingMode.INDIRECT);

			if (ArrayUtils.isNotEmpty(resourceIDs))
			{
				XResource resource = configurationController.getResource(resourceIDs[0]);
				pane = UnoRuntime.queryInterface(XPane.class, resource);
			}
		}
		return pane;
	}
	
	public static XResourceId [] getConfiguredResourceIDs(XComponent xComponent, XComponentContext xContext)
	{
		// Zugriff auf die Resourcen via XConfigurationController 
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XConfigurationController configurationController = xcontollerManager.getConfigurationController();
		XConfiguration configuration = configurationController.getCurrentConfiguration();
		
		XResourceId xAnchorId = ResourceId.createEmpty(xContext);
		XResourceId [] resourceIds = configuration.getResources(xAnchorId, "", AnchorBindingMode.INDIRECT);
		return resourceIds;
	}

	public static XPane getConfiguredPane(XComponent xComponent, XComponentContext xContext, XResourceId paneResourceID)
	{
		XPane pane = null;  
				
		// Zugriff auf die Resourcen via XConfigurationController 
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XConfigurationController configurationController = xcontollerManager.getConfigurationController();
		XConfiguration configuration = configurationController.getCurrentConfiguration();
	
		if(!paneResourceID.hasAnchor())
		{
			XResource resource = configurationController.getResource(paneResourceID);
			pane = UnoRuntime.queryInterface(XPane.class, resource);
		}
		
		return pane;
	}
	
	public static XView getConfiguredView(XComponent xComponent, XComponentContext xContext, XResourceId paneResourceID)
	{
		XView xView = null;  
				
		// Zugriff auf die Resourcen via XConfigurationController 
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XConfigurationController configurationController = xcontollerManager.getConfigurationController();
		XConfiguration configuration = configurationController.getCurrentConfiguration();
	
		if(paneResourceID.hasAnchor())
		{
			XResource resource = configurationController.getResource(paneResourceID);
			xView = UnoRuntime.queryInterface(XView.class, resource);
		}
		
		return xView;
	}

	public static XView getDrawXWindow(XComponent xComponent, XComponentContext xContext)
	{
		// Zugriff auf die Resourcen via XConfigurationController 
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XConfigurationController configurationController = xcontollerManager.getConfigurationController();
		XConfiguration configuration = configurationController.getCurrentConfiguration();
		
		XResourceId resourceIDempty = ResourceId.createEmpty(xContext);
		XResourceId[] resourceIDs = configuration.getResources(
				resourceIDempty, drawingFrameworkPanes[0],AnchorBindingMode.DIRECT);
		resourceIDs = configuration.getResources(resourceIDs[0],
				drawingFrameworkViews [1],AnchorBindingMode.DIRECT);
		XResource resource = configurationController.getResource(resourceIDs[0]);
		XView xView = UnoRuntime.queryInterface(XView.class, resource);
		return xView;
	}
	
	public static XWindow getCentrePaneWindow1(XComponent xComponent, XComponentContext xContext)
	{
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XConfigurationController configurationController = xcontollerManager.getConfigurationController();
		XConfiguration configuration = configurationController.getCurrentConfiguration();
		
		XResourceId resourceIDempty = ResourceId.createEmpty(xContext);
		XResourceId[] resourceIDs = configuration.getResources(
				resourceIDempty, drawingFrameworkPanes[2],AnchorBindingMode.DIRECT);
		
		XResource resource = configurationController.getResource(resourceIDs[0]);
		XPane pane = UnoRuntime.queryInterface(XPane.class, resource);
		
		/*
		resourceIDs = configuration.getResources(resourceIDs[0],
				drawingFrameworkViews [1],AnchorBindingMode.DIRECT);
		XResource resource = configurationController.getResource(resourceIDs[0]);
		XView xView = UnoRuntime.queryInterface(XView.class, resource);
		XPane pane = UnoRuntime.queryInterface(XPane.class,
				configurationController
						.getResource(resourceIDs[0].getAnchor()));
						
						*/
		
		return pane.getWindow();
	}
	
	public static XPane getCentrePane(XComponent xComponent, XComponentContext xContext)
	{
		// Zugriff auf die Resourcen via XConfigurationController 
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XControllerManager xcontollerManager = UnoRuntime.queryInterface(XControllerManager.class, xController);
		XConfigurationController configurationController = xcontollerManager.getConfigurationController();
		XConfiguration configuration = configurationController.getCurrentConfiguration();
		XResourceId resourceIDempty = ResourceId.createEmpty(xContext);
		XResourceId[] resourceIDs = configuration.getResources(
				resourceIDempty, drawingFrameworkPanes[2],AnchorBindingMode.DIRECT);
		
		XResource resource = configurationController.getResource(resourceIDs[0]);
		XPane pane = UnoRuntime.queryInterface(XPane.class, resource);
		return pane;
	}
	
	



	/*
	 * 
	 *  Page
	 * 
	 */

	public static Size getPageSize(XDrawPage xDrawPage)
	{
		Integer width = (Integer) getDrawPagePropertyValue(xDrawPage, "Width");
		Integer height = (Integer) getDrawPagePropertyValue(xDrawPage, "Height");
		Size size = new Size(width, height);
		return size;
	}

	public static Point getPageLeftTopBorder(XDrawPage xDrawPage)
	{
		Integer borderLeft = (Integer) getDrawPagePropertyValue(xDrawPage, "BorderLeft");
		Integer bordertop = (Integer) getDrawPagePropertyValue(xDrawPage, "BorderTop");		
		return new Point(borderLeft, bordertop);
	}

	public static int getPageBorderLeft(XDrawPage xDrawPage)
	{
		Integer borderLeft = (Integer) getDrawPagePropertyValue(xDrawPage, "BorderLeft");		
		return borderLeft.intValue();
	}
	
	public static Object getDrawPagePropertyValue(XDrawPage xDrawPage, String propertyName)
	{
		XPropertySet xPageProperties = UnoRuntime.queryInterface(XPropertySet.class, xDrawPage);		
		try
		{
			Object obj = xPageProperties.getPropertyValue(propertyName);
			return obj;
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

	/**
	 * Eine neue Seite hinzufuegen.
	 * 
	 * @param xComponent
	 * @param pageName
	 */
	static public void addPage(XComponent xComponent, String pageName)
	{		
		int n = getDrawPageCount(xComponent);			
		
		XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime
				.queryInterface(XDrawPagesSupplier.class, xComponent);
		XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
		XDrawPage xDrawPage = xDrawPages.insertNewByIndex( n);
		
		XNamed xNamed = UnoRuntime.queryInterface(XNamed.class, xDrawPage);
		xNamed.setName(pageName);	
	}
	
	/**
	 * Gibt den Namen der DrawPage zurueck. 
	 * LibrOffice konvertiert intern systemeigen Pagenamen 'page' in Locale Namen z.B. 'Folie' um.
	 * Diese Funktion simuliert diese Konvertierung. 
	 * 
	 * @param xDrawPage
	 * @return
	 */
	static public String getPageName(XDrawPage xDrawPage)
	{
		// localPageNamesMap initialisieren, da static
		initLocalPageNameMap();
		
		XNamed xNamed = UnoRuntime.queryInterface(XNamed.class, xDrawPage);
		String name = xNamed.getName();

		String localName = localPageNamesMap.keySet().iterator().next();
		if (StringUtils.startsWith(name, localName))
			 name = StringUtils.replace(name, localName, localPageNamesMap.get(localName));

		return name;
	}
	
	/**
	 * Die XDrawPage mit dem Namen 'pageName' zurueckgeben:
	 * 
	 * @param xComponent
	 * @param pageName
	 * @return
	 */
	static public XDrawPage getPage(XComponent xComponent, String pageName)
	{
		XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime
				.queryInterface(XDrawPagesSupplier.class, xComponent);

		XDrawPage xDrawPage;
		int count = getDrawPageCount(xComponent);
		for (int i = 0; i < count; i++)
		{
			xDrawPage = getDrawPageByIndex(xComponent, i);
			XPropertySet xPageProperties = UnoRuntime
					.queryInterface(XPropertySet.class, xDrawPage);

			String name = getPageName(xDrawPage);

			// String name = (String)
			// xPageProperties.getPropertyValue("LinkDisplayName");
			if (StringUtils.equals(name, pageName))
				return getDrawPageByIndex(xComponent, i);
		}
		return null;
	}
	
    /**
     * @param xComponent
     * @param currentPage
     */
    static public void setCurrentPage(XComponent xComponent, XDrawPage xDrawPage)
	{
		try
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
			XController xController = xModel.getCurrentController();
			XPropertySet xPageProperties = UnoRuntime
					.queryInterface(XPropertySet.class, xController);
			xPageProperties.setPropertyValue("CurrentPage", xDrawPage);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
    
    static public XDrawPage getCurrentPage(XComponent xComponent)
 	{
 		try
 		{
 			XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
 			XController xController = xModel.getCurrentController();
 			XPropertySet xPageProperties = UnoRuntime
 					.queryInterface(XPropertySet.class, xController);
 			Any any = (Any) xPageProperties.getPropertyValue("CurrentPage");
 			return(UnoRuntime.queryInterface(XDrawPage.class, any));
 			
 		} catch (Exception e)
 		{
 			// TODO Auto-generated catch block
 			e.printStackTrace();
 		}
 		return null;
 	}


	/**
	 * Alle DrawPage in einer Liste zurueckgeben.
	 * 
	 */
	public static List<XDrawPage> getDrawPages(XComponent xComponent)
	{
		List<XDrawPage>pageList = new ArrayList<XDrawPage>();
		int pageCount = getDrawPageCount(xComponent);
		
		for(int i = 0; i < pageCount; i++)
			pageList.add(getDrawPageByIndex(xComponent, i));
		
		return pageList;
	}
	
	/**
	 * Eine Page durch Index zurueckgeben.
	 * 
	 * @param xComponent
	 * @param nIndex
	 * @return
	 */
	public static XDrawPage getDrawPageByIndex(XComponent xComponent, int nIndex)
	{
		XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent);
		XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();

		try
		{
			return UnoRuntime.queryInterface(XDrawPage.class,xDrawPages.getByIndex(nIndex));
		} catch (IndexOutOfBoundsException e)
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
	
	/**
	 * Anzahl der Seiten zurueckgeben.
	 * 
	 * @param xComponent
	 * @return
	 */
	public static int getDrawPageCount(XComponent xComponent)
	{
		XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime
				.queryInterface(XDrawPagesSupplier.class, xComponent);
		XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
		return xDrawPages.getCount();
	}

	/**
	 * Den Namen der aktuellen Page zurueckgeben
	 * 
	 * @param xComponent
	 * @return
	 */
	public static String getCurrentPageName(XComponent xComponent)
	{
		try
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
 			XController xController = xModel.getCurrentController();
 			XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xController);
			Object objPage = xPropertySet.getPropertyValue("CurrentPage");
			if(objPage instanceof Any)
			{		
				Any any = (Any) objPage;
				XDrawPage xDrawPage = UnoRuntime.queryInterface(XDrawPage.class, any);				
				XNamed xNamed = UnoRuntime.queryInterface(XNamed.class, xDrawPage);
		        return xNamed.getName();      				
			}

		} catch (UnknownPropertyException e1)
		{
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (WrappedTargetException e1)
		{
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}			
		
		return null;
	}
	
	/*
	 * 
	 *  Layer
	 * 
	 */
	
	/**
	 * Alle Layernamen in einer Liste zurueckgeben.
	 * @param xComponent
	 * @param local = Umwandlung in Local (Umwandlung der Namen)
	 * @return
	 */
	public static List<String> readLayer(XComponent xComponent, boolean local)
	{
		List<String>layerList = new ArrayList<String>();
		
		if(xComponent != null)
		{
			XLayerManager xLayerManager = null;
			XLayerSupplier xLayerSupplier = UnoRuntime.queryInterface(
					XLayerSupplier.class, xComponent);
			XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
			xLayerManager = UnoRuntime.queryInterface(XLayerManager.class, xNameAccess );
			
			// alle Layernamen
			XNameAccess nameAccess = (XNameAccess) UnoRuntime.queryInterface( 
					XNameAccess.class, xLayerManager);			
			String [] names = nameAccess.getElementNames();
			if(ArrayUtils.isNotEmpty(names))
				for(String name : names)
				{
					if(local)
					{
						String localName = getLocalLayerName(name);
						if(localName != null)
						{
							if(StringUtils.isEmpty(localName))
								continue;
							name = localName;
						}
					}
					
					layerList.add(name);
				}				
		}
		
		return layerList;
	}
	
	
	/**
	 * Einen Layer selektiern.
	 * 
	 * @param xComponent
	 * @param pageName
	 * @param layerName
	 */
	public static void selectLayer(XComponent xComponent, String layerName)
	{
		XLayer xLayer = findLayer(xComponent, layerName);
		selectLayer(xComponent, xLayer);
	}

	/**
	 * Einen Layer selektiern.
	 * 
	 * @param xComponent
	 * @param xLayer
	 */
	public static void selectLayer(XComponent xComponent, XLayer xLayer)
	{		
		if (xLayer != null)
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
			XController xController = xModel.getCurrentController();

			XPropertySet xPageProperties = UnoRuntime
					.queryInterface(XPropertySet.class, xController);
			
			try
			{
				xPageProperties.setPropertyValue("ActiveLayer", xLayer);
			} catch (IllegalArgumentException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (UnknownPropertyException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (PropertyVetoException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WrappedTargetException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	/**
	 * Einen Layer selektiern.
	 * 
	 * @param xComponent
	 * @param pageName
	 * @param layerName
	 */
	public static XLayer getCurrentLayer(XComponent xComponent)
	{
		XLayer xLayer = null;
		try
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
			XController xController = xModel.getCurrentController();
			XPropertySet props = UnoRuntime
					.queryInterface(XPropertySet.class, xController); 
			Any any = (Any) props.getPropertyValue("ActiveLayer");
			xLayer = UnoRuntime.queryInterface(XLayer.class, any);	
			
		} catch (UnknownPropertyException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return xLayer;		
	}

	/**
	 * Gibt den Namen des Layers zurueck in dem 'xShape' definiert ist.
	 * 
	 * @param xComponent
	 * @param pageName
	 * @param xShape
	 * @return
	 */
	public static XLayer getLayerforShape(XComponent xComponent, XShape xShape)
	{
		XLayer layer = null;
		
		XLayerManager xLayerManager = null;
		XLayerSupplier xLayerSupplier = UnoRuntime.queryInterface(
				XLayerSupplier.class, xComponent);
		XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
		xLayerManager = UnoRuntime.queryInterface(XLayerManager.class, xNameAccess );		
		layer = xLayerManager.getLayerForShape(xShape);
		return layer;
	}

	
	/**
	 * Gibt den Namen des Layers zurueck in dem 'xShape' definiert ist.
	 * 
	 * @param xComponent
	 * @param pageName
	 * @param xShape
	 * @return
	 */
	public static String findLayer(XComponent xComponent, String pageName, XShape xShape)
	{
		List<String>layerNames = readLayer(xComponent, false);
		if(!layerNames.isEmpty())
		{
			for(String layerName : layerNames)
			{
				List<XShape>layerShapes = getLayerShapes(xComponent, pageName, layerName);
				if(layerShapes.contains(xShape))
					return layerName;				
			}
		}
		
		return null;
	}
	
	/**
	 * Einen Layer ueber seinen Namen suchen. Erwartet wird der LocalName.
	 * 
	 * @param xComponent
	 * @param pageName
	 * @param layerName
	 * @return
	 */
	public static XLayer findLayer(XComponent xComponent, String layerName)
	{
		try
		{
			XLayerManager xLayerManager = null;
			XLayerSupplier xLayerSupplier = UnoRuntime.queryInterface(
					XLayerSupplier.class, xComponent);
			XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
			xLayerManager = UnoRuntime.queryInterface(XLayerManager.class, xNameAccess );
			
			// alle Layernamen
			XNameAccess nameAccess = (XNameAccess) UnoRuntime.queryInterface( 
					XNameAccess.class, xLayerManager);	
			
			String [] names = nameAccess.getElementNames();		
			if (ArrayUtils.isNotEmpty(names))
			{
				for (String name : names)
				{
					// layweName umwandeln in systemdefinierten Namen
					for(String localLayerName : localLayerNamesMap.keySet())
					{
						String checkLocalName = localLayerNamesMap.get(localLayerName);
						if(StringUtils.equals(checkLocalName, layerName))
						{
							layerName = localLayerName;
							break;
						}
					}
					
					Any any = (Any) xNameAccess.getByName(name);
					XLayer xLayer = (XLayer) UnoRuntime.queryInterface(XLayer.class, any);
					
					if (StringUtils.equals(name, layerName))
						return xLayer;
				}
			}
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		
		return null;
	}
	
	public static void setLockedState(XLayer xLayer, boolean status)
	{
		try
		{		
			XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, xLayer); 
			props.setPropertyValue("IsLocked", status);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	/*
	 * 
	 *  Shapes
	 * 
	 * 
	 */
	
	/**
	 * Die Shapes eines Layers auflisten.
	 * 
	 * @param xComponent
	 * @param pageName
	 * @param layerName
	 * @return
	 */
	public static List<XShape> getLayerShapes(XComponent xComponent, String pageName, String layerName)
	{
		List<XShape> shapeList = new ArrayList<XShape>();

		try
		{
			XDrawPage drawPage = PageHelper.getDrawPageByName(xComponent,pageName);
			if (drawPage != null)
			{
				XShapes xShapes = UnoRuntime.queryInterface(XShapes.class,drawPage);
				if (xShapes != null)
				{
					int n = xShapes.getCount();
					for (int i = 0; i < n; i++)
					{
						Object obj = xShapes.getByIndex(i);
						if (obj instanceof Any)
						{
							Any any = (Any) xShapes.getByIndex(i);
							XShape xShape = UnoRuntime
									.queryInterface(XShape.class, any);
							if (xShape != null)
							{
								XPropertySet xPropSet = (XPropertySet) UnoRuntime
										.queryInterface(XPropertySet.class,xShape);
								String name = (String) xPropSet.getPropertyValue("LayerName");
								if (StringUtils.equals(layerName, name))
								{
									if (xShape != null)
										shapeList.add(xShape);
								}
							}
						}
					}
				}
			}

		} catch (IndexOutOfBoundsException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (UnknownPropertyException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return shapeList;
	}
	
	/**
	 * Alle Shapes der Seite 'pageName' in einer Liste zurueckgeben
	 *  
	 * @param xComponent
	 * @param pageName
	 * @return
	 */
	public static List<XShape> getSelectedShapes(XComponent xComponent) throws DisposedException
	{
		Any any;
		List<XShape>shapeList = new ArrayList<XShape>();
		
		try
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
			XController xController = xModel.getCurrentController();
			XSelectionSupplier selectionSupplier = UnoRuntime
					.queryInterface(XSelectionSupplier.class, xController);
			Object selObj = selectionSupplier.getSelection();
			if (selObj instanceof Any)
			{
				any = (Any) selObj;
				XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, any);
				if (xShapes != null)
				{
					int n = xShapes.getCount();

					// die Shapes einer Liste sammeln
					for (int i = 0; i < n; i++)
					{
						Object obj = xShapes.getByIndex(i);
						if (obj instanceof Any)
						{
							any = (Any) xShapes.getByIndex(i);
							XShape xShape = UnoRuntime.queryInterface(XShape.class, any);
							
							Point pt = xShape.getPosition();
							System.out.println(pt.X+"  "+pt.Y);
							
							
							if (xShape != null)
								shapeList.add(xShape);
						}
					}
				}
			}

		} catch (IndexOutOfBoundsException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return shapeList;
	}
	
	/**
	 * Alle Shapes der Seite 'pageName' in einer Liste zurueckgeben
	 *  
	 * @param xComponent
	 * @param pageName
	 * @return
	 */
	public static List<XShape> readPageShapes(XComponent xComponent, String pageName)
	{		
		List<XShape>shapeList = new ArrayList<XShape>();
		try
		{
			// Shapes einer Page ermitteln
			XDrawPage drawPage = PageHelper.getDrawPageByName(xComponent, pageName);
			XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, drawPage);	
			int n = xShapes.getCount();
			
			// die Shapes einer Liste sammeln
			for(int i = 0;i < n;i++)
			{
				Object obj = xShapes.getByIndex(i);
				if(obj instanceof Any)
				{
					Any any = (Any) xShapes.getByIndex(i);
					XShape xShape = UnoRuntime.queryInterface(XShape.class, any);
					if(xShape != null)
						shapeList.add(xShape);			
				}
			}
		} catch (IndexOutOfBoundsException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return shapeList;
	}

	/**
	 * Den Label des Shapes (z.Formen, Linien etc) zurueckgeben.
	 * 
	 * @param xShape
	 * @return
	 */
	public static String getShapeLabel(XShape xShape)
	{		
		String label = null;
		try
		{
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
			label =  (String) xPropSet.getPropertyValue("UINamePlural");
		} catch (UnknownPropertyException | WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return label;			
	}
	
	public static XShape createShape(XComponent xComponent, Point pos, Size size, String shapeType)
	{
		XShape xShape = null;
		try
		{
			xShape = createInstanceMSF(xComponent, XShape.class, shapeType);
			xShape.setPosition( pos );
			xShape.setSize( size );
		} catch (PropertyVetoException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return xShape;
	}
}
