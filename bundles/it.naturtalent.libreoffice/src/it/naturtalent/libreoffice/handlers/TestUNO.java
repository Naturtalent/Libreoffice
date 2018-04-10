 
package it.naturtalent.libreoffice.handlers;

import java.util.List;

import it.naturtalent.libreoffice.Bootstrap;
import it.naturtalent.libreoffice.PageHelper;
import it.naturtalent.libreoffice.Utils;
import it.naturtalent.libreoffice.draw.DrawDocument;

import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Status;
import org.eclipse.core.runtime.jobs.Job;
import org.eclipse.e4.core.di.annotations.Execute;

import com.sun.star.awt.Size;
import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XMap;
import com.sun.star.container.XNameAccess;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XLayer;
import com.sun.star.drawing.XLayerManager;
import com.sun.star.drawing.XMasterPageTarget;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.Any;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;


/**
 * Nur zum propieren.
 * 
 * 
 * 
 * @author A682055
 *
 */
public class TestUNO
{
	
	private static final String TESTURL="/media/dieter/f8ceb1a1-74b6-4dbf-a487-e12e6249ced01/home/dieter/temp/Grundstücksmasse.odg";
	
	private static String WIN_TESTURL = "C:\\Users\\A682055\\Daten\\Naturtalent4\\workspaces\\telekom\\1410350687331-1\\Strukturpanung\\planung.odg";
		
	
	private static final String WIN_TESTURL_STICK="E:\\planung.odg";
	private static final String LINUX_TESTURL_STICK="/media/dieter/TOSHIBA/planung.odg";
	
	
	@Execute
	public void execute()
	{
		System.out.println("UNO");
		firstStep();
	}
	
	private void firstStep()
	{

		XComponentContext xContext = null;

		try
		{
			
			DrawDocument drawDocument = new DrawDocument();
			drawDocument.loadPage(WIN_TESTURL);
			//drawDocument.loadPage(WIN_TESTURL_STICK);
			
			//DrawDocument.load(WIN_TESTURL);
			
			
			//XComponent xComponent = Utils.newDocComponentFromTemplate(WIN_TESTURL);
			
			//XComponent xComponent = Utils.newDocComponent(Utils.OfficeTypes.DRAW.getType());
			
			
			//loadComponent(LINUX_TESTURL_STICK);
			
			
			/*
			XMultiServiceFactory xFactory =
	                UnoRuntime.queryInterface(
	                        XMultiServiceFactory.class, xComponent );
			
			String [] names = xFactory.getAvailableServiceNames();
			
			for(String name : names)
				System.out.println(name);
				*/

				
			
			
			
			//DocumentSettings = (Settings) xFactory.createInstance("com.sun.star.drawing.DocumentSettings");
			
			
			/*
			XMultiServiceFactory msf = (XMultiServiceFactory)
					xServiceManager.createInstanceWithContext("com.sun.star.lang.XMultiServiceFactory",xContext);
					*/
			
			
			//Settings docSettings = (Settings) msf.createInstance(" com.sun.star.document.Settings");
			
			
			//System.out.println(docSettings);
			
			
			
			/*
			XComponent xComponent = Utils.newDocComponent(Utils.OfficeTypes.DRAW.getType());
			System.out.println(xComponent);

			List<XPropertySet>layers = Utils.getLayerPropertySet(xComponent);			
			for(XPropertySet layer : layers)
			{
				System.out.println(layer.getPropertyValue(Utils.LayerProperties.IsVisible.name()));
			}
			*/
			
			
			
			
		} catch (Exception e)
		{
			e.printStackTrace(System.err);
			System.exit(1);
		}
	}
	
	private void loadComponent(final String url)
	{
		Job j = new Job("Load Job") //$NON-NLS-1$
		{
			@Override
			protected IStatus run(final IProgressMonitor monitor)
			{
				try
				{
					XComponent xComponent = Utils.newDocComponentFromTemplate(url);

					
					XMultiServiceFactory xFactory =
			                UnoRuntime.queryInterface(
			                        XMultiServiceFactory.class, xComponent );
					XInterface settings = (XInterface) xFactory.createInstance("com.sun.star.drawing.DocumentSettings");					
					XPropertySet xPageProperties = UnoRuntime.queryInterface( XPropertySet.class, settings );					
					Integer scaleNumerator = (Integer) xPageProperties.getPropertyValue("ScaleNumerator");
					Integer scaleDenominator = (Integer) xPageProperties.getPropertyValue("ScaleDenominator");

					System.out.println("Maßstab: "+scaleNumerator+":"+scaleDenominator);
					
					scaleDenominator = 100;
					xPageProperties.setPropertyValue("ScaleDenominator", scaleDenominator);
					System.out.println("Maßstab: "+scaleNumerator+":"+scaleDenominator);
					
					Short measureUnit = (Short) xPageProperties.getPropertyValue("MeasureUnit");
					System.out.println("MeasureUnit: "+scaleNumerator+":"+measureUnit);
					measureUnit = 2;
					xPageProperties.setPropertyValue("MeasureUnit", measureUnit);
					 
					
					/*
					XDrawPage xDrawPage = getDrawPage(xComponent); 
					
					Size size = PageHelper.getPageSize(xDrawPage);
					XPropertySet xPageProperties = UnoRuntime.queryInterface( XPropertySet.class, xDrawPage );
					XPropertySetInfo propInfo = xPageProperties.getPropertySetInfo();
					
					Property [] properties = propInfo.getProperties();
					for(Property property : properties)
					{
						System.out.println(property.Name);
					}
					*/
					
					
					/*
					 XServiceInfo xInfo = UnoRuntime.queryInterface(
				                XServiceInfo.class, xComponent );				
					 String [] names = xInfo.getSupportedServiceNames();
					 for(String name : names)
					 {
						 System.out.println(name);
					 }
					 */


					 /*
					XPropertySet xProperySet = (XPropertySet) UnoRuntime
							.queryInterface(XPropertySet.class,
									xComponent);
					XPropertySetInfo propInfo = xProperySet.getPropertySetInfo();
					Property [] properties = propInfo.getProperties();
					for(Property property : properties)
					{
						System.out.println(property.Name);
					}
					*/
					
					
					System.out.println("geladen");
					
					
				} catch (Exception e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				return Status.OK_STATUS;
			}
		};
		j.schedule();	
	}
	
	private void exampleSDraw(XComponentContext xContext)
	{
		// oooooooooooooooooooooooooooStep
		// 2oooooooooooooooooooooooooooooooooooooooo
		// open an empty document. In this case it's a draw document.
		// For this purpose an instance of com.sun.star.frame.Desktop
		// is created. It's interface XDesktop provides the XComponentLoader,
		// which is used to open the document via loadComponentFromURL

		com.sun.star.lang.XComponent xDrawDoc = null;
		System.out.println("Opening an empty Draw document ...");
		xDrawDoc = openDraw(xContext);
		
		// listener wird aufgerufen ueber die Fuktion XComponent.dispose();
		xDrawDoc.addEventListener(new XEventListener()
		{			
			@Override
			public void disposing(EventObject arg0) 
			{
				System.out.println("ende");
				// TODO Auto-generated method stub
				
			}
		});
		
		// oooooooooooooooooooooooooooStep
		// 3oooooooooooooooooooooooooooooooooooooooo
		// get the drawpage an insert some shapes.
		// the documents DrawPageSupplier supplies the DrawPage vi IndexAccess
		// To add a shape get the MultiServiceFaktory of the document, create an
		// instance of the ShapeType and add it to the Shapes-container
		// provided by the drawpage
		XDrawPage xDrawPage =  getDrawPage(xDrawDoc); 		
		System.out.println(xDrawPage);
		
		
		XShape sequence = createDrawingSequence(xDrawDoc, xDrawPage); 
		
		// put something on the drawpage
		System.out.println("inserting some Shapes");
		com.sun.star.drawing.XShapes xShapes = UnoRuntime.queryInterface(
				com.sun.star.drawing.XShapes.class, xDrawPage);
		xShapes.add(createShape(xDrawDoc, 2000, 1500, 1000, 1000, "Line", 0));
		xShapes.add(createShape(xDrawDoc, 3000, 4500, 15000, 1000, "Ellipse",
				16711680));
		xShapes.add(createShape(xDrawDoc, 5000, 3500, 7500, 5000, "Rectangle",
				6710932));

		System.out.println("done");
		System.exit(0);

	}

	public static com.sun.star.lang.XComponent openDraw(
			com.sun.star.uno.XComponentContext xContext)
	{
		com.sun.star.frame.XComponentLoader xCLoader;
		com.sun.star.text.XTextDocument xDoc = null;
		com.sun.star.lang.XComponent xComp = null;
	
		try
		{
			// get the remote office service manager
			com.sun.star.lang.XMultiComponentFactory xMCF = xContext
					.getServiceManager();
	
			Object oDesktop = xMCF.createInstanceWithContext(
					"com.sun.star.frame.Desktop", xContext);
	
			xCLoader = UnoRuntime.queryInterface(
					com.sun.star.frame.XComponentLoader.class, oDesktop);
			com.sun.star.beans.PropertyValue szEmptyArgs[] = new com.sun.star.beans.PropertyValue[0];
			String strDoc = "private:factory/sdraw";
			xComp = xCLoader.loadComponentFromURL(strDoc, "_blank", 0,
					szEmptyArgs);
	
		} catch (Exception e)
		{
			System.err.println(" Exception " + e);
			e.printStackTrace(System.err);
		}
	
		return xComp;
	}
	
	// get the drawpage of drawing here
	private XDrawPage getDrawPage(XComponent xDrawDoc) 
	{
		XDrawPage xDrawPage = null;
		
		try
		{
			System.out.println("getting Drawpage");
			com.sun.star.drawing.XDrawPagesSupplier xDPS = UnoRuntime
					.queryInterface(
							com.sun.star.drawing.XDrawPagesSupplier.class,xDrawDoc);
			com.sun.star.drawing.XDrawPages xDPn = xDPS.getDrawPages();
			com.sun.star.container.XIndexAccess xDPi = UnoRuntime
					.queryInterface(com.sun.star.container.XIndexAccess.class,
							xDPn);
			xDrawPage = UnoRuntime.queryInterface(
					com.sun.star.drawing.XDrawPage.class, xDPi.getByIndex(0));
		} catch (Exception e)
		{
			System.err.println("Couldn't create document" + e);
			e.printStackTrace(System.err);
		}

		return xDrawPage;
	}
	
	public static com.sun.star.drawing.XShape createShape(
			com.sun.star.lang.XComponent xDocComp, int height, int width,
			int x, int y, String kind, int col)
	{
		// possible values for kind are 'Ellipse', 'Line' and 'Rectangle'
		com.sun.star.awt.Size size = new com.sun.star.awt.Size();
		com.sun.star.awt.Point position = new com.sun.star.awt.Point();
		com.sun.star.drawing.XShape xShape = null;

		// get MSF
		com.sun.star.lang.XMultiServiceFactory xDocMSF = UnoRuntime
				.queryInterface(com.sun.star.lang.XMultiServiceFactory.class,
						xDocComp);

		try
		{
			Object oInt = xDocMSF.createInstance("com.sun.star.drawing." + kind
					+ "Shape");
			xShape = UnoRuntime.queryInterface(
					com.sun.star.drawing.XShape.class, oInt);
			size.Height = height;
			size.Width = width;
			position.X = x;
			position.Y = y;
			xShape.setSize(size);
			xShape.setPosition(position);

		} catch (Exception e)
		{
			System.err.println("Couldn't create instance " + e);
			e.printStackTrace(System.err);
		}

		com.sun.star.beans.XPropertySet xSPS = UnoRuntime.queryInterface(
				com.sun.star.beans.XPropertySet.class, xShape);

		try
		{
			xSPS.setPropertyValue("FillColor", new Integer(col));
		} catch (Exception e)
		{
			System.err.println("Can't change colors " + e);
			e.printStackTrace(System.err);
		}

		return xShape;
	}

	
	public static com.sun.star.drawing.XShape createDrawingSequence(
			com.sun.star.lang.XComponent xDocComp,
			com.sun.star.drawing.XDrawPage xDP)
	{
		com.sun.star.awt.Size size = new com.sun.star.awt.Size();
		com.sun.star.awt.Point position = new com.sun.star.awt.Point();
		com.sun.star.drawing.XShape xShape = null;
		com.sun.star.drawing.XShapes xShapes = UnoRuntime.queryInterface(
				com.sun.star.drawing.XShapes.class, xDP);
		int height = 3000;
		int width = 3500;
		int x = 1900;
		int y = 20000;
		Object oInt = null;
		int r = 40;
		int g = 0;
		int b = 80;

		// get MSF
		com.sun.star.lang.XMultiServiceFactory xDocMSF = UnoRuntime
				.queryInterface(com.sun.star.lang.XMultiServiceFactory.class,
						xDocComp);

		for (int i = 0; i < 370; i = i + 25)
		{
			try
			{
				oInt = xDocMSF
						.createInstance("com.sun.star.drawing.EllipseShape");
				xShape = UnoRuntime.queryInterface(
						com.sun.star.drawing.XShape.class, oInt);
				size.Height = height;
				size.Width = width;
				position.X = (x + (i * 40));
				position.Y = (new Float(y + (Math.sin((i * Math.PI) / 180))
						* 5000)).intValue();
				xShape.setSize(size);
				xShape.setPosition(position);

			} catch (Exception e)
			{
				// Some exception occurs.FAILED
				System.err.println("Couldn't get Shape " + e);
				e.printStackTrace(System.err);
			}

			b = b + 8;

			com.sun.star.beans.XPropertySet xSPS = UnoRuntime.queryInterface(
					com.sun.star.beans.XPropertySet.class, xShape);

			try
			{
				xSPS.setPropertyValue("FillColor", new Integer(getCol(r, g, b)));
				xSPS.setPropertyValue("Shadow", new Boolean(true));
			} catch (Exception e)
			{
				System.err.println("Can't change colors " + e);
				e.printStackTrace(System.err);
			}
			xShapes.add(xShape);
		}

		com.sun.star.drawing.XShapeGrouper xSGrouper = UnoRuntime
				.queryInterface(com.sun.star.drawing.XShapeGrouper.class, xDP);

		xShape = xSGrouper.group(xShapes);

		return xShape;
	}

	public static int getCol(int r, int g, int b)
	{
		return r * 65536 + g * 256 + b;
	}


	/*
	 *  Writer 
	 * 
	 * 
	 * 
	 */
	
	private void exampleSWriter(XComponentContext xContext)
	{
		XTextDocument openWriter = openWriter(xContext);
		System.out.println(xContext);
		
	}
	
	public static XTextDocument openWriter(XComponentContext xContext)
	{
		// define variables
		com.sun.star.frame.XComponentLoader xCLoader;
		com.sun.star.text.XTextDocument xDoc = null;
		com.sun.star.lang.XComponent xComp = null;

		try
		{
			// get the remote office service manager
			com.sun.star.lang.XMultiComponentFactory xMCF = xContext
					.getServiceManager();

			Object oDesktop = xMCF.createInstanceWithContext(
					"com.sun.star.frame.Desktop", xContext);

			xCLoader = UnoRuntime.queryInterface(
					com.sun.star.frame.XComponentLoader.class, oDesktop);
			com.sun.star.beans.PropertyValue[] szEmptyArgs = new com.sun.star.beans.PropertyValue[0];
			String strDoc = "private:factory/swriter";
			xComp = xCLoader.loadComponentFromURL(strDoc, "_blank", 0,
					szEmptyArgs);
			xDoc = UnoRuntime.queryInterface(
					com.sun.star.text.XTextDocument.class, xComp);

		} catch (Exception e)
		{
			System.err.println(" Exception " + e);
			e.printStackTrace(System.err);
		}
		return xDoc;
	}
	
}