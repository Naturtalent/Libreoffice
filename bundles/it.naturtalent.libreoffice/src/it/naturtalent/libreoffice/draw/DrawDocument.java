package it.naturtalent.libreoffice.draw;

import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Random;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Status;
import org.eclipse.core.runtime.jobs.Job;
import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.internal.workbench.E4Workbench;
import org.eclipse.e4.ui.model.application.MApplication;
import org.eclipse.e4.ui.workbench.IWorkbench;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.swt.graphics.Rectangle;
import org.eclipse.swt.widgets.Display;
import org.jnativehook.GlobalScreen;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.accessibility.XAccessibleComponent;
import com.sun.star.accessibility.XAccessibleContext;
import com.sun.star.awt.Point;
//import com.sun.star.awt.Rectangle;
import com.sun.star.awt.Size;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyChangeEvent;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertyChangeListener;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.HomogenMatrixLine3;
import com.sun.star.drawing.PolyPolygonBezierCoords;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XLayer;
import com.sun.star.drawing.XLayerManager;
import com.sun.star.drawing.XLayerSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.Any;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;
import com.sun.star.view.XSelectionSupplier;

import it.naturtalent.libreoffice.Bootstrap;
import it.naturtalent.libreoffice.DesignHelper;
import it.naturtalent.libreoffice.DrawDocumentEvent;
import it.naturtalent.libreoffice.DrawDocumentUtils;
import it.naturtalent.libreoffice.DrawPagePropertyListener;
import it.naturtalent.libreoffice.DrawShape;
import it.naturtalent.libreoffice.DrawShape.SHAPEPROP;
import it.naturtalent.libreoffice.FrameActionListener;
import it.naturtalent.libreoffice.PageHelper;
//import it.naturtalent.libreoffice.ServiceManager;
import it.naturtalent.libreoffice.ShapeSelectionListener;
import it.naturtalent.libreoffice.Utils;
import it.naturtalent.libreoffice.listeners.GlobalMouseListener;
import it.naturtalent.libreoffice.utils.Props;

public class DrawDocument
{
	protected XComponent xComponent;

	private XComponentContext xContext;

	private XDesktop xDesktop;

	private XFrame xFrame;

	private XSelectionSupplier selectionSupplier;

	private static boolean atWork = false;

	// die aktuelle Seite
	// protected int drawPageIndex = 0;
	// protected XDrawPage drawPage;

	protected String documentPath;

	// Name der aktuellen Seite
	protected String pageName;

	protected IEventBroker eventBroker;

	// Massstab und Masseinheit
	protected Scale scale;

	// registriert die Layer des Documents
	private Map<String, Layer> layerRegistry = new HashMap<String, Layer>();

	// der zuletzt selektierte Layer
	private Layer lastSelectedLayer;

	public PolyPolygonBezierCoords aCoords;

	// Listener informiert, wenn DrawDocument extern (LibreOffice) geschlossen
	// wurde
	private TerminateListener terminateListener;

	// Map<TerminateListener, DrawPagePath> (@see TerminateListener)
	public static Map<TerminateListener, DrawDocument> openTerminateDocumentMap = new HashMap<TerminateListener, DrawDocument>();

	private DrawPagePropertyListener drawPagePropertyListener;

	private ShapeSelectionListener shapeSelectionListener;

	private FrameActionListener frameActionListener;

	private Log log = LogFactory.getLog(this.getClass());

	// der Globale MouseListener
	private GlobalMouseListener globalMouseListener;

	// TopWindow der des Draw-Documents
	private XTopWindow xTopWindow;

	// Stempelmodus ein-/ausschalten
	private boolean stampMode = false;

	private XWindow xComponentWindow;

	private XWindow xContainerWindow;

	//private Map<XLayer, ILayerLayout> layoutMap = new HashMap<XLayer, ILayerLayout>();

	//private ILayerLayout currentLayout;

	/**
	 * 
	 * 
	 * 
	 * 
	 */
	public DrawDocument()
	{
		MApplication currentApplication = E4Workbench.getServiceContext()
				.get(IWorkbench.class).getApplication();
		eventBroker = currentApplication.getContext().get(IEventBroker.class);
	}

	public void loadPage(final String documentPath)
	{
		final Job j = new Job("Load Job") //$NON-NLS-1$
		{
			@Override
			protected IStatus run(final IProgressMonitor monitor)
			{
				try
				{
					loadDocument(documentPath);
					// setDocumentProperties();
				} catch (Exception e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				return Status.OK_STATUS;
			}
		};

		/*
		 * j.addJobChangeListener(new JobChangeAdapter() {
		 * 
		 * @Override public void done(IJobChangeEvent event) {
		 * if(!event.getResult().isOK()) { // Fehler (wahrscheinlich keine
		 * 'JPIPE' - LibraryPath) IStatus status = event.getResult(); final
		 * String message = status.toString();
		 * 
		 * Display.getDefault().syncExec(new Runnable() { public void run() { //
		 * Watchdog (@see OpenDesignAction) abschalten
		 * eventBroker.post(DrawDocumentEvent.
		 * DRAWDOCUMENT_EVENT_DOCUMENT_OPEN_CANCEL, null);
		 * 
		 * MessageDialog .openError( Display.getDefault().getActiveShell(),
		 * "Error Load Dokument",message); j.cancel(); } }); }
		 * super.done(event); }
		 * 
		 * });
		 */

		j.schedule();
	}

	private void loadDocument(String documentPath) throws Exception
	{
		this.documentPath = documentPath;

		File sourceFile = new java.io.File(documentPath);
		StringBuffer sTemplateFileUrl = new StringBuffer("file:///");
		sTemplateFileUrl
				.append(sourceFile.getCanonicalPath().replace('\\', '/'));

		xContext = Bootstrap.bootstrap();
		if (xContext != null)
		{
			XMultiComponentFactory xMCF = xContext.getServiceManager();

			// retrieve the Desktop object, we need its XComponentLoader
			Object desktop = xMCF.createInstanceWithContext(
					"com.sun.star.frame.Desktop", xContext);

			XComponentLoader xComponentLoader = UnoRuntime
					.queryInterface(XComponentLoader.class, desktop);

			//
			//
			//

			/*
			 * Object svgfilter; try { svgfilter =
			 * xMCF.createInstanceWithContext(
			 * "com.sun.star.document.SVGFilter", xComponentLoader); XFilter
			 * xfilter = (XFilter) UnoRuntime.queryInterface( XFilter.class,
			 * svgfilter ); XImporter ximporter = (XImporter)
			 * UnoRuntime.queryInterface( XExporter.class, svgfilter );
			 * System.out.println("Inserting image...");
			 * 
			 * } catch (com.sun.star.uno.Exception e) { // TODO Auto-generated
			 * catch block e.printStackTrace(); }
			 */

			//
			//
			//

			// load
			PropertyValue[] loadProps = new PropertyValue[0];
			xComponent = xComponentLoader.loadComponentFromURL(
					sTemplateFileUrl.toString(), "_blank", 0, loadProps);

			// empirisch ermittelt
			Thread.sleep(500);

			xComponent.addEventListener(new XEventListener()
			{
				@Override
				public void disposing(EventObject arg0)
				{
					// EventBroker informiert ueber das Schliessen des Dokuments
					eventBroker.post(
							DrawDocumentEvent.DRAWDOCUMENT_EVENT_DOCUMENT_CLOSE,
							DrawDocument.this);
				}
			});

			// TerminateListener (registriert eine durch Libreoffice ausgeloeste
			// Close-Aktionen)
			terminateListener = new TerminateListener();
			terminateListener.setEventBroker(eventBroker);
			xDesktop = UnoRuntime.queryInterface(XDesktop.class, desktop);
			xDesktop.addTerminateListener(terminateListener);

			// EventBroker informiert, dass Ladevorgang abgeschlossen ist
			eventBroker.post(DrawDocumentEvent.DRAWDOCUMENT_EVENT_DOCUMENT_OPEN,
					this);

			// PageListener installieren und aktivierten
			drawPagePropertyListener = new DrawPagePropertyListener(xComponent);
			drawPagePropertyListener.activatePageListener();

			/*
			 * Shapeselection Listener
			 */
			XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
			XController xController = xModel.getCurrentController();
			selectionSupplier = UnoRuntime
					.queryInterface(XSelectionSupplier.class, xController);
			shapeSelectionListener = new ShapeSelectionListener();
			selectionSupplier
					.addSelectionChangeListener(shapeSelectionListener);

			// Listener meldet Layeraenderung die vom Modelllayer ausgegangen
			// ist oder durch Selektion eines
			// Shapes in einem anderen Layer - (nicht durch direkte Tabselektion
			// im DrawDocument)
			XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class,
					xController);
			Props.showProps("XComponent", props);
			props.addPropertyChangeListener("ActiveLayer",
					new XPropertyChangeListener()
					{

						@Override
						public void disposing(EventObject arg0)
						{
							// TODO Auto-generated method stub

						}

						@Override
						public void propertyChange(PropertyChangeEvent arg0)
						{
							System.out.println("DrawDocument().315 - L A Y E R");

						}
					});
			// ActiveLayerPropertyListener layerListener = new
			// ActiveLayerPropertyListener(xComponent);
			// layerListener.activatePageListener();

			// Listener ueberwacht die Frameaktivitaeten (z.B. Frame wird
			// aktiviert)
			xFrame = xController.getFrame();
			frameActionListener = new FrameActionListener();
			xFrame.addFrameActionListener(frameActionListener);
			// das geoeffnete Dokument mit Listener als Key speichern
			openTerminateDocumentMap.put(terminateListener, this);

			/*
			 * -------------------------- experimental
			 * -----------------------------------
			 */

			// Globalen MouseListener einschalten und Logger begrenzen

			/*
			 * GlobalScreen.registerNativeHook(); globalMouseListener = new
			 * GlobalMouseListener(eventBroker); Logger logger =
			 * Logger.getLogger(GlobalScreen.class.getPackage().getName());
			 * //logger.setLevel(Level.OFF);
			 * 
			 * GlobalScreen.addNativeMouseListener(globalMouseListener);
			 * GlobalScreen.removeNativeMouseListener(globalMouseListener); try
			 * { GlobalScreen.unregisterNativeHook(); } catch
			 * (NativeHookException e1) { e1.printStackTrace(); }
			 * System.runFinalization();
			 * 
			 * 
			 * // den globalen MouseListener ein-/ausschalten xTopWindow =
			 * DrawDocumentUtils.getDrawDocumentXTopWindow(xContext);
			 * xTopWindow.addTopWindowListener(new XTopWindowAdapter() {
			 * 
			 * @Override public void windowDeactivated(EventObject arg0) { //
			 * TopWindow wurde deaktiviert - MouseListener ausschalten
			 * //System.out.println("DEACTIVATE");
			 * GlobalScreen.removeNativeMouseListener(globalMouseListener); }
			 * 
			 * @Override public void windowActivated(EventObject arg0) { //
			 * TopWindow wurde aktiviert - MouseListener einschalten
			 * //System.out.println("ACTIVATE");
			 * GlobalScreen.addNativeMouseListener(globalMouseListener); } });
			 */

			// die erste Seite wird selektioert
			List<XDrawPage> drawPages = DrawDocumentUtils
					.getDrawPages(xComponent);
			eventBroker.post(DrawDocumentEvent.DRAWDOCUMENT_PAGECHANGE_PROPERTY,
					drawPages.get(0));

			System.out.println("END load DrawDocument");
		}
	}

	public void doActivateShapeListener()
	{
		selectionSupplier.removeSelectionChangeListener(shapeSelectionListener);
	}

	/*
	 * Die Funktion wird u.a. getriggert durch die Selektion eines Shapes im
	 * DrawDocument. (@see ShapeSelectionListener) u. (@see
	 * handleShapeSelectedEvent())
	 * 
	 * !!! Getriggert wird dieser Event aber auch, wenn das DrawDocument extern
	 * geschlossen wurde. Hierdurch wird eine DisposedException ausgeloest da
	 * beim Zugriff durch 'DrawDocumentUtils.getSelectedShapes()' sich das
	 * XModel bereits im Zustand 'disposed' befindet.
	 * 
	 * !!! Moegliche Ursache UI DrawDodument ist zu diesem Zeitpungkt bereits
	 * geschlossen
	 * 
	 * Die Funktion sucht den Layer des markierten Shapes und selektiert den
	 * Layer.
	 * 
	 */
	public void doShapeSelection(Object arg0)
	{
		List<XShape> shapeList = DrawDocumentUtils
				.getSelectedShapes(xComponent);
		if ((shapeList != null) && (!shapeList.isEmpty()))
		{
			XLayer xLayer = DrawDocumentUtils.getLayerforShape(xComponent,
					shapeList.get(0));
			DrawDocumentUtils.selectLayer(xComponent, xLayer);
		}
	}
	
	public void removeShapeSelectionListener()
	{
		XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
		XController xController = xModel.getCurrentController();
		selectionSupplier = UnoRuntime
				.queryInterface(XSelectionSupplier.class, xController);
		shapeSelectionListener = new ShapeSelectionListener();
		selectionSupplier
				.removeSelectionChangeListener(shapeSelectionListener);
	}

	/**
	 * Seite mit dem Namen 'pageName' im DrawDocument selektieren.
	 * 
	 * @param pageName
	 */
	boolean flag = false; // experimetell ein-/ausschalten Window

	public void selectPage(String pageName)
	{
		XDrawPage xDrawPage = DrawDocumentUtils.getPage(xComponent, pageName);
		if (xDrawPage != null)
			DrawDocumentUtils.setCurrentPage(xComponent, xDrawPage);

		int borderLeft = DrawDocumentUtils.getPageBorderLeft(xDrawPage);

		System.out.println(borderLeft);
		// xComponentWindow.setVisible(flag);
		// flag = !flag;

	}

	/**
	 * Ueberpruefen und Anpassen der Mouseposition, die durch den globalen
	 * MouseListener zureckgegeben werden.
	 * 
	 * (Left/Top- Position des TopWindows (0,0))
	 * 
	 * Zurueckgegeben wird eine Position die an das TopWindow angepasst ist oder
	 * null wenn dieses nicht tangiert wird.
	 * 
	 */

	public Point containsPoint(int x, int y)
	{
		XAccessible xTopWindowAccessible = (XAccessible) UnoRuntime
				.queryInterface(XAccessible.class, xTopWindow);
		XAccessibleContext xAccessibleContext = xTopWindowAccessible
				.getAccessibleContext();
		XAccessibleComponent aComp = (XAccessibleComponent) UnoRuntime
				.queryInterface(XAccessibleComponent.class, xAccessibleContext);

		// pruefen, ob die vom globalen Listener gemeldete Mouseposition im
		// TopWindow des DrawDokuments liegt
		Point ptScreen = aComp.getLocationOnScreen();
		Size size = aComp.getSize();

		if ((x < ptScreen.X) || (x > (ptScreen.X + size.Width)))
			return null;

		if ((y < ptScreen.Y) || (y > (ptScreen.Y + size.Height)))
			return null;

		// globale MousePosition pt an TopWindow anpassen
		Point pt = new Point(x, y);
		pt.X = pt.X - ptScreen.X;
		pt.Y = pt.Y - ptScreen.Y;

		return pt;
	}

	/**
	 * Realisiert einen Globalen MouseClick.
	 * 
	 * @param mousePoint
	 */
	public void doGlobalMouseEvent(Object mousePoint)
	{
		Point pos = getStatusbarPosition();
		
		/*
		if ((pos != null) && (currentLayout != null))
		{

			// Testshape hinzufuegen
			Size size = new Size(1000, 1000);
			XShape xShape = DrawDocumentUtils.createShape(xComponent, pos, size,
					"com.sun.star.drawing.RectangleShape");
			XDrawPage xDrawPage = DrawDocumentUtils.getCurrentPage(xComponent);
			xDrawPage.add(xShape);

		}
		*/
	}
	
	/**
	 * Statusbarposition auslesen und transformieren um die Laengeneinheit und
	 * die aktuelle DrawBorderEinstellung.
	 * 
	 * @return
	 */
	public Point getStatusbarPosition()
	{
		Point pos = null;

		XAccessibleContext accessibleContext = DrawDocumentUtils
				.getAccessibleContext(xContext);
		Double[] statusPos = DrawDocumentUtils
				.getStatusposition(accessibleContext);
		if (statusPos != null)
		{
			// 'LeftTop' - Position der Page Borderdefinition
			XDrawPage xDrawPage = DrawDocumentUtils.getCurrentPage(xComponent);
			Point LeftTop = DrawDocumentUtils.getPageLeftTopBorder(xDrawPage);

			//
			// ToDo
			//
			// die MousePosition an Laengeneinheit anpassen
			//

			// eaperimentell
			statusPos[0] = statusPos[0] * 1000.0;
			statusPos[1] = statusPos[1] * 1000.0;

			// Postion in Point ueberfuehren
			pos = new Point(statusPos[0].intValue(), statusPos[1].intValue());

			// Mouseposition an Borderdefinition der Page anpassen
			pos.X = pos.X + LeftTop.X;
			pos.Y = pos.Y + LeftTop.Y;
		}
		else
		{
			MessageDialog.openInformation(Display.getDefault().getActiveShell(),
					"", "Statusbar einschalten");
		}

		return pos;
	}

	public boolean isStampMode()
	{
		return stampMode;
	}

	public void setStampMode(boolean stampMode)
	{
		this.stampMode = stampMode;
	}

	public boolean isLayerSelect()
	{

		return false;
	}

	/**
	 * Layer mit dme namen 'layerName' selektieren.
	 * 
	 * @param layerName
	 */
	public void selectLayer(String layerName)
	{
		DrawDocumentUtils.selectLayer(xComponent, layerName);
	}

	public Object getPage(String pageName)
	{
		return DrawDocumentUtils.getPage(xComponent, pageName);
	}

	public String getCurrentPage()
	{
		return PageHelper.getCurrentPage(xComponent);
	}

	/*
	 * 
	 *  
	 */
	public boolean isChildPageByFrame(Object xFrame)
	{
		return (xFrame.equals(xFrame));
	}

	/*
	 * Ueberprueft, ob die uebergebene Seite 'xDrawPage' zu diesem DrawDocument
	 * gehoert.
	 * 
	 */
	public boolean isChildPage(Object xDrawPage)
	{
		List<XDrawPage> pages = DrawDocumentUtils.getDrawPages(xComponent);
		return pages.contains(xDrawPage);
	}

	/**
	 * Setzt den Focus auf diese Zeichnung. Sind mehrere Zeichnungen geoffnet,
	 * wird diese Zeichnung bearbeitbar sichtbar im Desktop gezeigt.
	 */
	public void setFocus()
	{
		DesignHelper.setFocus(xComponent);
	}

	/**
	 * Die Namen aller Pages in einer Liste zurueckgeben
	 * 
	 * @return
	 */
	public List<String> getAllPages(boolean local)
	{
		List<String> allPages = new ArrayList<>();

		int count = PageHelper.getDrawPageCount(xComponent);
		for (int i = 0; i < count; i++)
		{
			XDrawPage page;
			try
			{
				page = PageHelper.getDrawPageByIndex(xComponent, i);
				if (local)
					allPages.add(DrawDocumentUtils.getPageName(page));
				else
				{
					XNamed xNamed = UnoRuntime.queryInterface(XNamed.class,page);
					allPages.add(xNamed.getName());
				}

			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return allPages;
	}
	
	/**
	 * Liest alle XShapes eines Layers.
	 * Fuer jedes xShape wird eine unspezifischen DrawShape erzeugt und diese werden in einer
	 * Liste zurueckgegeben.
	 * 
	 * @param pageName
	 * @param layerName
	 * @return
	 */
	public List<DrawShape>getLayerShapes(String pageName, String layerName)
	{
		List<DrawShape>drawShapes = new ArrayList<DrawShape>();
	
		List<XShape>xShapes = DrawDocumentUtils.getLayerShapes(xComponent, pageName, layerName);		
		for(XShape xShape : xShapes)
		{
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
			try
			{
				com.sun.star.awt.Rectangle rect = (com.sun.star.awt.Rectangle) xPropSet.getPropertyValue("BoundRect");
				DrawShape drawShape = new DrawShape(rect.X, rect.Y, rect.Width, rect.Height);
				drawShape.setDrawShapeType(SHAPEPROP.Label);
				drawShape.xShape = xShape;
				drawShapes.add(drawShape);

			} catch (UnknownPropertyException | WrappedTargetException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		return drawShapes;
	}
	
	/**
	 * Universelle DrawShape-Propertyabfrage
	 * 
	 * @param drawProperty - in DrawShape definiert
	 * @return
	 */
	public Object getDrawShapeProperty(DrawShape drawShape)
	{
		switch (drawShape.getDrawShapeType())
			{
				case Label:
					return /* String */ DrawDocumentUtils.getShapeLabel((XShape) (drawShape.xShape));
										
				case Line:
					
					ShapeLineLengthUtils lengthUtils = new ShapeLineLengthUtils();					
					return /* BigDecimal */ lengthUtils.getLineLength((XShape) (drawShape.xShape));

				default:
					break;
			}
		
		return null;
	}

	/**
	 * Die gelisteten Shapes 'drawShapes' werden in die aktuelle Seite eigefuegt und somit sichtbar.
	 * Ist der 'layerName' != null, werden die Shapes diesem Layer zugeordnet.
	 * Ist der 'layerName' == null, werden die Shapes dem Layer 'layout' zugeordnet. 
	 * 
	 * @param layerName
	 * @param drawShapes
	 */
	public void setLayerShapes(String layerName, List<DrawShape>drawShapes)
	{
		XShape xShape;
		XLayer xLayer = null;
		XDrawPage xDrawPage = DrawDocumentUtils.getCurrentPage(xComponent);
		if(StringUtils.isNotEmpty(layerName))
			xLayer = DrawDocumentUtils.findLayer(xComponent, layerName);

		if(xDrawPage != null)
		{
			// xShapes = Shapesliste der aktuellen Page zu der die 'drawShapes' hinzugefuegt werden
			XShapes xShapes = UnoRuntime.queryInterface(XShapes.class,xDrawPage);
			for (DrawShape drawShape : drawShapes)
			{
				Point pos = new Point(drawShape.getX(), drawShape.getY());
				Size size = new Size(drawShape.getWidht(),drawShape.getHeight());
				String xType = null;
				switch (drawShape.getDrawShapeType())
					{
						// der Linientyp wird durch Analyse der xShape Daten
						// spezifiziert
						case Line:
							if (drawShape.xShape == null)
							{
								xType = DrawDocumentUtils.LineShapeType;
								break;
							}

						default:
							break;
					}
				
				if (xType != null)
				{
					// xShape erzeugen und den 'page'-Shapes hinzufuegen  
					xShape = DrawDocumentUtils.createShape(xComponent, pos,size, xType);					
					xShapes.add(xShape);
					
					// den erzeugten xShape im DrawShape speichern 
					drawShape.setxShape(xShape);
					
					// ist ein xLayer definiert, werden die Shapes diesem zugeordnet
					if(xLayer != null)
						DrawDocumentUtils.attacheShapeToLayer(xComponent, xLayer, xShape);
				}
			}
		}
	}
	
	/**
	 * Shapes in der aktuellen Seite loeschen.
	 * 
	 * @param drawShapes
	 */
	public void removeShapes(List<DrawShape>drawShapes)
	{
		XDrawPage xDrawPage = DrawDocumentUtils.getCurrentPage(xComponent);
		if(xDrawPage != null)
		{
			XShapes xShapes = UnoRuntime.queryInterface(XShapes.class,xDrawPage);
			for (DrawShape drawShape : drawShapes)		
				xShapes.remove((XShape) drawShape.getxShape());		
		}
	}
	
	public void selectShape(DrawShape drawShape)
	{
		DrawDocumentUtils.shapeSelection(xComponent, (XShape) drawShape.xShape);
	}
	
	/**
	 * @param propertyName
	 * @param value
	 */
	public void setDocumentSettings(String propertyName, Object value)
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
				xPageProperties.setPropertyValue(propertyName, value);
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}	
	}
	
	/**
	 * @param propertyName
	 * @return
	 */
	public Object getDocumentStettings(String propertyName)
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
				return xPageProperties.getPropertyValue(propertyName);
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
			
		return null;
	}
	
	/**
	 * Layerproperties setzen
	 * Die Propertynamen sind in DrawDocumentUtils definiert.
	 * 
	 * @param propertyName
	 * @param value
	 */
	public void setLayerProperty(String layerName, String propertyName, Object value)
	{
		if(StringUtils.equals(propertyName, DrawDocumentUtils.LAYERLOCK))
		{
			XLayer xLayer = DrawDocumentUtils.findLayer(xComponent, layerName);
			if(xLayer != null)
			{
				DrawDocumentUtils.setLayerLock(xLayer, (boolean) value);
				return;
			}
		}
	}
	
	/**
	 * Layerproperties abfragen
	 * Die Propertynamen sind in DrawDocumentUtils definiert.
	 * 
	 * @param layerName
	 * @param propertyName
	 * @return
	 */
	public Object getLayerProperty(String layerName, String propertyName)
	{
		if(StringUtils.equals(propertyName, DrawDocumentUtils.LAYERLOCK))
		{
			XLayer xLayer = DrawDocumentUtils.findLayer(xComponent, layerName);
			if(xLayer != null)
				return DrawDocumentUtils.getLayerLock(xLayer);
		}
		return null;
	}

		
	/**
	 * Die Shapes eines Layers zurueckgeben
	 * 
	 * @return
	 */
	public List<String> getAllLayers(boolean local)
	{
		return DrawDocumentUtils.readLayer(xComponent, local);
	}

	public void closeDesktop()
	{
		xDesktop.terminate();
	}

	/**
	 * Schliesst DrawDocument ausgelöst durch eine CloseAktion
	 * (Kontext-/Toolaction). Close durch LibroOffice-Aktion wird hier nicht
	 * registriert, @see it.naturtalent.libreoffice.draw.TerminateListener
	 */
	public void closeDocument()
	{
		// global MouseListener deaktivieren
		GlobalScreen.removeNativeMouseListener(globalMouseListener);

		// verursacht Exception: Ursache unklar
		// com.sun.star.lang.DisposedException: java_remote_bridge
		// com.sun.star.lib.uno.bridges.java_remote.java_remote_bridge@1ca8eaf
		// is disposed
		// !!! Moegliche Ursache UI DrawDodument ist zu diesem Zeitpungkt
		// bereits geschlossen
		xComponent.dispose();
	}

	/**
	 * Die aktuelle Skalierung einlesen
	 */
	public void readScaleData()
	{
		scale = new Scale(xComponent);
		scale.pullScaleProperties();
	}

	public void scalePoint(int x, int y)
	{
		Point pt = new Point(x, y);
		pt = scale.scalePoint(pt);
		System.out.println("Scale: " + pt.X + " | " + pt.Y);
	}

	/*
	 * public void pullScaleData() { scale = new Scale(xComponent);
	 * scale.pullScaleProperties(); }
	 */

	public Scale getScale()
	{
		return scale;
	}

	public Rectangle getPageBound()
	{
		try
		{

			XDrawPage xdrawPage = PageHelper.getDrawPageByName(xComponent,
					pageName);
			Size aPageSize = PageHelper.getPageSize(xdrawPage);

			// Size aPageSize = PageHelper.getPageSize(drawPage);

			int nHalfWidth = aPageSize.Width / 2;
			int nHalfHeight = aPageSize.Height / 2;

			Random aRndGen = new Random();
			int nRndObjWidth = aRndGen.nextInt(nHalfWidth);
			int nRndObjHeight = aRndGen.nextInt(nHalfHeight);

			int nRndObjPosX = aRndGen.nextInt(nHalfWidth - nRndObjWidth)
					+ nRndObjWidth;
			int nRndObjPosY = aRndGen.nextInt(nHalfHeight - nRndObjHeight)
					+ nHalfHeight;

			return new Rectangle(nRndObjPosX, nRndObjPosY, nRndObjWidth,
					nRndObjHeight);

		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	private void pullPageSettings() throws Exception
	{
		if (xContext != null)
		{
			XMultiServiceFactory xFactory = UnoRuntime
					.queryInterface(XMultiServiceFactory.class, xComponent);
			XInterface settings = (XInterface) xFactory
					.createInstance("com.sun.star.drawing.DocumentSettings");
			XPropertySet xPageProperties = UnoRuntime
					.queryInterface(XPropertySet.class, settings);
			Integer scaleNumerator = (Integer) xPageProperties
					.getPropertyValue("ScaleNumerator");
			Integer scaleDenominator = (Integer) xPageProperties
					.getPropertyValue("ScaleDenominator");

			System.out.println(
					"Maßstab: " + scaleNumerator + ":" + scaleDenominator);

			scaleDenominator = 100;
			xPageProperties.setPropertyValue("ScaleDenominator",
					scaleDenominator);
			System.out.println(
					"Maßstab: " + scaleNumerator + ":" + scaleDenominator);

			Short measureUnit = (Short) xPageProperties
					.getPropertyValue("MeasureUnit");
			System.out.println("MeasureUnit: " + measureUnit);
			measureUnit = 2;
			xPageProperties.setPropertyValue("MeasureUnit", measureUnit);
		}

	}

	public static boolean isAtWork()
	{
		return atWork;
	}

	public XComponent getxComponent()
	{
		return xComponent;
	}

	public XDesktop getXDesktop()
	{
		return xDesktop;
	}

	public XComponentContext getxContext()
	{
		return xContext;
	}

	public XDrawPage getXDrawPage()
	{
		return PageHelper.getDrawPageByName(xComponent, pageName);
	}

	public XFrame getXframe()
	{
		return xFrame;
	}

	/**
	 * Die aktuelle Seite einstellen einstellen
	 * 
	 */

	/*
	 * public void setDrawPage(int pageIndex) { try { drawPage =
	 * PageHelper.getDrawPageByIndex(xComponent, pageIndex);
	 * 
	 * } catch (Exception e) { // TODO: handle exception e.printStackTrace(); }
	 * }
	 */

	public String getDocumentPath()
	{
		return documentPath;
	}

	/**
	 * Die Seite mit dem Namen 'pageName' aktivieren. Sollte diese Seite nicht
	 * existieren, wird eine Neue erzeugt und an oberster Position eingefuegt.
	 * 
	 * @param pageName
	 */
	public void selectDrawPage(String pageName)
	{
		XDrawPage xDrawPage = PageHelper.getDrawPageByName(xComponent,
				pageName);
		if (xDrawPage == null)
		{
			try
			{
				// neue Seite mit dem Namen 'pageName' einfuegen
				xDrawPage = PageHelper.insertNewDrawPageByIndex(xComponent, 0);
				XPropertySet xPageProperties = UnoRuntime
						.queryInterface(XPropertySet.class, xDrawPage);

				PageHelper.setPageName(xDrawPage, pageName);

			} catch (Exception e)
			{
				e.printStackTrace();
			}
		}

		// Seite akkivieren
		if (xDrawPage != null)
			PageHelper.setCurrentPage(xComponent, xDrawPage);
	}

	public void selectDrawLayer(Layer layer)
	{
		// zunaechst alle Layer deaktivieren
		Layer[] allLayers = getLayers();
		for (Layer deactivelayer : allLayers)
		{
			deactivelayer.deactivate();
			deactivelayer.setLocked(true);
		}

		// Ziellayer aktivieren
		layer.activate();
		layer.setLocked(false);

		// den selektierten Layer speichern
		lastSelectedLayer = layer;
	}

	public void addDrawPage(String pageName)
	{
		DrawDocumentUtils.addPage(xComponent, pageName);

		/*
		 * XDrawPage xDrawPage; try { xDrawPage =
		 * PageHelper.insertNewDrawPageByIndex(xComponent, 0); if(xDrawPage !=
		 * null) PageHelper.setPageName(xDrawPage, pageName);
		 * 
		 * } catch (Exception e) { }
		 */
	}

	public String readPageName(Object drawPage)
	{
		return (drawPage instanceof XDrawPage)
				? DrawDocumentUtils.getPageName((XDrawPage) drawPage)
				: null;
	}

	public void removeDrawPage(String pageName)
	{
		XDrawPage xDrawPage = PageHelper.getDrawPageByName(xComponent,
				pageName);
		if (xDrawPage != null)
			PageHelper.removeDrawPage(xComponent, xDrawPage);
	}

	public void renameDrawPage(String pageName, String newName)
	{
		XDrawPage xDrawPage = PageHelper.getDrawPageByName(xComponent,
				pageName);
		if (xDrawPage != null)
			PageHelper.setPageName(xDrawPage, newName);
	}

	/**
	 * Alle Layer der Seite einlesen
	 * 
	 */

	/*
	 * public void pullStyle() { if(xComponent != null) { Style graphicStyle =
	 * new Style(xComponent); Integer linecolor = graphicStyle.getLineColor();
	 * 
	 * System.out.println(Integer.toHexString(linecolor));
	 * 
	 * //graphicStyle.setLineColor(new Integer( 0xff0000 ));
	 * 
	 * //Style graphicStyle = new Style(xComponent); //XStyle style =
	 * graphicStyle.getStyle("LineStyle");
	 * //System.out.println(style.getName()); } }
	 */

	public void pullStyleOLD()
	{
		if (xComponent != null)
		{
			try
			{
				// Graphics Style Container
				XModel xModel = UnoRuntime.queryInterface(XModel.class,
						xComponent);
				com.sun.star.style.XStyleFamiliesSupplier xSFS = UnoRuntime
						.queryInterface(
								com.sun.star.style.XStyleFamiliesSupplier.class,
								xModel);
				com.sun.star.container.XNameAccess xFamilies = xSFS
						.getStyleFamilies();

				String[] Families = xFamilies.getElementNames();
				for (int i = 0; i < Families.length; i++)
				{
					// this is the family
					System.out.println("\n" + Families[i]);

					// and now all available styles
					Object aFamilyObj = xFamilies.getByName(Families[i]);
					com.sun.star.container.XNameAccess xStyles = UnoRuntime
							.queryInterface(
									com.sun.star.container.XNameAccess.class,
									aFamilyObj);
					String[] Styles = xStyles.getElementNames();
					for (int j = 0; j < Styles.length; j++)
					{
						System.out.println("   " + Styles[j]);
						Object aStyleObj = xStyles.getByName(Styles[j]);
						com.sun.star.style.XStyle xStyle = UnoRuntime
								.queryInterface(com.sun.star.style.XStyle.class,
										aStyleObj);
						// now we have the XStyle Interface and the CharColor
						// for
						// all styles is exemplary be set to red.
						XPropertySet xStylePropSet = UnoRuntime
								.queryInterface(XPropertySet.class, xStyle);

						if (xStylePropSet != null)
							Utils.printPropertyNames(xStylePropSet);

						/*
						 * XPropertySetInfo xStylePropSetInfo =
						 * xStylePropSet.getPropertySetInfo(); if (
						 * xStylePropSetInfo.hasPropertyByName( "CharColor" ) )
						 * { xStylePropSet.setPropertyValue( "CharColor", new
						 * Integer( 0xff0000 ) ); }
						 */
					}
				}
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	/**
	 * Alle Layer der Seite einlesen
	 * 
	 */
	public void pullLayer()
	{
		if (xComponent != null)
		{
			try
			{
				XLayerManager xLayerManager = null;
				XLayerSupplier xLayerSupplier = UnoRuntime
						.queryInterface(XLayerSupplier.class, xComponent);
				XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
				xLayerManager = UnoRuntime.queryInterface(XLayerManager.class,
						xNameAccess);

				// alle Layernamen
				XNameAccess nameAccess = (XNameAccess) UnoRuntime
						.queryInterface(XNameAccess.class, xLayerManager);
				String[] names = nameAccess.getElementNames();

				layerRegistry.clear();
				for (String name : names)
				{
					Any any = (Any) xNameAccess.getByName(name);
					XLayer xLayer = (XLayer) UnoRuntime
							.queryInterface(XLayer.class, any);
					Layer layer = new Layer(xLayer);
					layer.setDrawDocument(this);
					layerRegistry.put(name, layer);
				}
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	/**
	 * einen neuen Layer hinzufuegen
	 * 
	 */
	public void addLayer(String layerName)
	{
		if (xComponent != null)
		{
			try
			{
				if (!layerRegistry.containsKey(layerName))
				{
					XLayerManager xLayerManager = null;
					XLayerSupplier xLayerSupplier = UnoRuntime
							.queryInterface(XLayerSupplier.class, xComponent);
					XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
					xLayerManager = UnoRuntime
							.queryInterface(XLayerManager.class, xNameAccess);

					// create a second layer
					XLayer xLayer = xLayerManager
							.insertNewByIndex(xLayerManager.getCount());

					Layer newLayer = new Layer(xLayer);
					newLayer.setDrawDocument(this);
					newLayer.setName(layerName);
					newLayer.setVisible(true);
					newLayer.setLocked(true);
					layerRegistry.put(layerName, newLayer);
				}

			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	public void visibleLayer(Layer layer)
	{
		// zunaechst alle Layer unsichtbar
		Layer[] allLayers = getLayers();
		for (Layer deactivelayer : allLayers)
			deactivelayer.setVisible(false);

		// Ziellayer sichtbar
		layer.setVisible(true);
	}

	public String getPageName()
	{
		return pageName;
	}

	public void setPageName(String pageName)
	{
		this.pageName = pageName;
	}

	public Layer getLastSelectedLayer()
	{
		return lastSelectedLayer;
	}

	/**
	 * Einen bestimmten Layer zuruekgeben.
	 * 
	 * @param layerName
	 * @return
	 */
	public Layer getLayer(String layerName)
	{
		return layerRegistry.get(layerName);
	}

	/**
	 * Alle Layer dieser Seite in einem Array zurueckgeben.
	 * 
	 * @return
	 */
	public List<String> getLayerNames()
	{
		List<String> listLayernames = new ArrayList<String>();
		listLayernames.addAll(layerRegistry.keySet());
		return listLayernames;
	}

	/**
	 * Alle Layer dieser Seite in einem Array zurueckgeben.
	 * 
	 * @return
	 */
	public Layer[] getLayers()
	{
		Layer[] layers = null;
		for (Iterator<Layer> itLayer = layerRegistry.values()
				.iterator(); itLayer.hasNext();)
			layers = ArrayUtils.add(layers, itLayer.next());
		return layers;
	}

	public Integer pullScaleDenominator()
	{
		if (xComponent != null)
		{
			try
			{
				XMultiServiceFactory xFactory = UnoRuntime
						.queryInterface(XMultiServiceFactory.class, xComponent);
				XInterface settings = (XInterface) xFactory.createInstance(
						"com.sun.star.drawing.DocumentSettings");
				XPropertySet xPageProperties = UnoRuntime
						.queryInterface(XPropertySet.class, settings);
				Integer scaleDenominator = (Integer) xPageProperties
						.getPropertyValue("ScaleDenominator");

				return scaleDenominator;

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

	/*
	 * public Double getScaleFactor() { if (xComponent != null) { try {
	 * XMultiServiceFactory xFactory = UnoRuntime.queryInterface(
	 * XMultiServiceFactory.class, xComponent); XInterface settings =
	 * (XInterface) xFactory
	 * .createInstance("com.sun.star.drawing.DocumentSettings"); XPropertySet
	 * xPageProperties = UnoRuntime.queryInterface( XPropertySet.class,
	 * settings); Integer scaleNumerator = (Integer) xPageProperties
	 * .getPropertyValue("ScaleNumerator"); Integer scaleDenominator = (Integer)
	 * xPageProperties .getPropertyValue("ScaleDenominator");
	 * 
	 * double scaleFactor = ((double)scaleDenominator/(double)scaleNumerator);
	 * return scaleFactor;
	 * 
	 * } catch (UnknownPropertyException e) { // TODO Auto-generated catch block
	 * e.printStackTrace(); } catch (WrappedTargetException e) { // TODO
	 * Auto-generated catch block e.printStackTrace(); } catch
	 * (com.sun.star.uno.Exception e) { // TODO Auto-generated catch block
	 * e.printStackTrace(); } }
	 * 
	 * return null; }
	 */

	public HomogenMatrixLine3 getScaleFactors()
	{
		if (xComponent != null)
		{
			try
			{
				HomogenMatrixLine3 scaleFactors = new HomogenMatrixLine3();

				XMultiServiceFactory xFactory = UnoRuntime
						.queryInterface(XMultiServiceFactory.class, xComponent);
				XInterface settings = (XInterface) xFactory.createInstance(
						"com.sun.star.drawing.DocumentSettings");
				XPropertySet xPageProperties = UnoRuntime
						.queryInterface(XPropertySet.class, settings);
				Integer scaleNumerator = (Integer) xPageProperties
						.getPropertyValue("ScaleNumerator");
				Integer scaleDenominator = (Integer) xPageProperties
						.getPropertyValue("ScaleDenominator");

				scaleFactors.Column1 = (double) scaleNumerator;
				scaleFactors.Column2 = (double) scaleDenominator;
				return scaleFactors;

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

	public Integer getZoomFaktor()
	{
		try
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
			XController xController = xModel.getCurrentController();

			XPropertySet xPageProperties = UnoRuntime
					.queryInterface(XPropertySet.class, xController);
			return new Integer(
					(int) xPageProperties.getPropertyValue("ZoomValue"));

		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

	public void setMeasureUnit(Short measureUnit)
	{
		try
		{
			XPropertySet xDrawDocumentProperties = getDrawDocumentProperties();
			xDrawDocumentProperties.setPropertyValue("MeasureUnit",
					measureUnit);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private XPropertySet getDrawDocumentProperties()
	{
		try
		{
			XMultiServiceFactory xFactory = UnoRuntime
					.queryInterface(XMultiServiceFactory.class, xComponent);
			XInterface settings = (XInterface) xFactory
					.createInstance("com.sun.star.drawing.DocumentSettings");
			XPropertySet xProperties = UnoRuntime
					.queryInterface(XPropertySet.class, settings);

			return xProperties;

		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}
	


}
