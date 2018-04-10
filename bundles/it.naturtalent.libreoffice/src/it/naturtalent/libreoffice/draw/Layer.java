package it.naturtalent.libreoffice.draw;



import it.naturtalent.libreoffice.PageHelper;
import it.naturtalent.libreoffice.Utils;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.eclipse.swt.graphics.Rectangle;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XLayer;
import com.sun.star.drawing.XLayerManager;
import com.sun.star.drawing.XLayerSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.Any;
import com.sun.star.uno.UnoRuntime;

public class Layer
{
	
	public enum DefaultLayerNames
	{
		  Masslinien, Layout, Steuerelemente;

		  public String getType()
		  {
		    if ( this == Masslinien)
		      return "measurelines";
		    
		    if ( this == Layout)
			      return "layout";
		    
		    if ( this == Steuerelemente)
			      return "controls";

		    return null;
		  }	
	}
	
	
	public static final String UNO_NAME_PROPERTY = "Name"; 
	public static final String UNO_VISIBLE_PROPERTY = "IsVisible";
	public static final String UNO_LOCK_PROPERTY = "IsLocked";
	public static final String UNO_DESCRIPTION_PROPERTY = "Description";
	
	public XLayer xLayer;
	private XPropertySet xLayerPropSet;	
	private DrawDocument drawDocument;	
	private Style style;
	
	
	
	
	public Layer(XLayer xLayer)
	{
		super();
		this.xLayer = xLayer;
		xLayerPropSet = UnoRuntime.queryInterface(XPropertySet.class, xLayer);
	}
	
	public DrawDocument getDrawDocument()
	{
		return drawDocument;
	}
	
	/**
	 * Layer wird durch das DrawDokument annektiert.
	 *  
	 * @param drawDocument
	 */
	public void setDrawDocument(DrawDocument drawDocument)
	{
		this.drawDocument = drawDocument;		
		initStyle();
	}

	public Style getStyle()
	{
		return style;
	}

	public void setStyle(Style style)
	{
		this.style = style;
	}


	public String getName()
	{
		if (xLayerPropSet != null)
		{
			try
			{
				return (String) xLayerPropSet.getPropertyValue(UNO_NAME_PROPERTY);
			} catch (UnknownPropertyException | WrappedTargetException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return null;
	}
	
	public void setName(String name)
	{
		if (xLayerPropSet != null)
		{
			try
			{
				xLayerPropSet.setPropertyValue(UNO_NAME_PROPERTY, name);
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	public boolean isVisible()
	{
		try
		{
			return (boolean) xLayerPropSet.getPropertyValue(UNO_VISIBLE_PROPERTY);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return false;
	}
	
	public void setVisible(boolean visible)
	{
		if (xLayerPropSet != null)
		{
			try
			{
				xLayerPropSet.setPropertyValue(UNO_VISIBLE_PROPERTY, visible);
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
		
	public boolean isLocked()
	{
		try
		{
			return (boolean) xLayerPropSet.getPropertyValue(UNO_LOCK_PROPERTY);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return false;
	}
	
	public void setLocked(boolean lock)
	{
		try
		{			
			xLayerPropSet.setPropertyValue(UNO_LOCK_PROPERTY, lock);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}		

	public void initStyle()
	{				
		if((drawDocument != null) && (drawDocument.xComponent != null))
		{
			style = new Style();
			
			// Styledaten einlesen
			style.pullLineColor(drawDocument);
		}
	}
	
	public boolean isActive()
	{				
		if((drawDocument != null) && (drawDocument.xComponent != null))
		{
			XComponent xComponent = drawDocument.xComponent;
			XPropertySet xPropertySet = getPropertgSet(xComponent);
			try
			{
				Any any = (Any) xPropertySet.getPropertyValue("ActiveLayer");
				XLayer activeLayer = UnoRuntime.queryInterface(XLayer.class, any);
				if(activeLayer != null)
				{
					String activeName = (String) activeLayer.getPropertyValue(UNO_NAME_PROPERTY);
					return StringUtils.equals(activeName, getName());
				}
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}			
		}
		
		return false;
	}

	
	public void deactivate()
	{				
		if((drawDocument != null) && (drawDocument.xComponent != null))
		{
			if (isActive())
			{
				XComponent xComponent = drawDocument.xComponent;

				// Styledaten einlesen
				style.pullLineColor(drawDocument);
			}
		}
	}

	public void activate()
	{				
		if((drawDocument != null) && (drawDocument.xComponent != null))
		{
			XComponent xComponent = drawDocument.xComponent;
			
			// Styledaten ausgeben
			style.pushLineColor(drawDocument);
			
			// Layer selektieren
			select(xComponent);
		}
	}
	
	private void select(XComponent xComponent)
	{
		if(xComponent != null)
		{
			 try
			{
				XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
				XController xController = xModel.getCurrentController();

				XPropertySet xPageProperties = UnoRuntime.queryInterface(
						XPropertySet.class, xController);
				xPageProperties.setPropertyValue("ActiveLayer", xLayer);
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}		
	}
	
	private XPropertySet getPropertgSet(XComponent xComponent)
	{
		try
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
			XController xController = xModel.getCurrentController();

			XPropertySet xPageProperties = UnoRuntime.queryInterface(
					XPropertySet.class, xController);
			return xPageProperties;			
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}
	
	public void addShape(Shape shape)
	{
		// the fact that the shape must have been added to the page before
        // it is possible to apply changes to the PropertySet, it is a good
        // proceeding to add the shape as soon as possible
		//XDrawPage drawPage = drawDocument.getDrawPage();
		XDrawPage drawPage = drawDocument.getXDrawPage();
        XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, drawPage);	    
        xShapes.add(shape.xShape);

		// Shape dem Layer zuordnen
        XLayerSupplier xLayerSupplier = UnoRuntime.queryInterface(
                XLayerSupplier.class, drawDocument.xComponent);
        XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
        XLayerManager xLayerManager = UnoRuntime.queryInterface(
                XLayerManager.class, xNameAccess );
        xLayerManager.attachShapeToLayer(shape.xShape, xLayer);       
	}
	
	public void deleteShapes()
	{
		//XDrawPage drawPage = drawDocument.getDrawPage();
		XDrawPage drawPage = drawDocument.getXDrawPage();
        XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, drawPage);	    

        List<Shape>shapes = pullLayerShapes();
        for(Shape shape : shapes)
        {
        	XShape xShape = shape.xShape;
        	xShapes.remove(xShape);
        }
	}
	
	
	public void deleteShape(Shape shape)
	{
		// Shape im Design loeschen
		//XDrawPage drawPage = drawDocument.getDrawPage();
		XDrawPage drawPage = drawDocument.getXDrawPage();
        XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, drawPage);	    
        xShapes.remove(shape.xShape);
	}

	/**
	 * 
	 * 
	 * @param shape
	 */
	public void pushShape(Shape shape)
	{
		// Shape zeichnen
		XDrawPage drawPage = drawDocument.getXDrawPage();
        XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, drawPage);	    
        xShapes.add(shape.xShape);
        
		// Shape dem Layer zuordnen
        XLayerSupplier xLayerSupplier = UnoRuntime.queryInterface(
                XLayerSupplier.class, drawDocument.xComponent);
        XNameAccess xNameAccess = xLayerSupplier.getLayerManager();
        XLayerManager xLayerManager = UnoRuntime.queryInterface(
                XLayerManager.class, xNameAccess );
        xLayerManager.attachShapeToLayer(shape.xShape, xLayer);
	}
	
	/**
	 * Liest alle definierten Shapes aus dem Design.
	 * 
	 * @return
	 */
	public List<Shape> pullLayerShapes()
	{
		 List<Shape>shapes = null;		
		if(drawDocument != null)
		{
			try
			{		
				shapes = new ArrayList<Shape>();
				XDrawPage drawPage = PageHelper.getDrawPageByName(drawDocument.xComponent, drawDocument.pageName);
				String myName = getName();
	            XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, drawPage);	            
	            int n = xShapes.getCount();
	            for(int i = 0;i < n;i++)
	            {
	            	Object obj = xShapes.getByIndex(i);
	            	if(obj instanceof Any)
	            	{
	            		Any any = (Any) xShapes.getByIndex(i);
	            		XShape xShape = UnoRuntime.queryInterface(XShape.class, any);
	            		if(xShape != null)
	            		{
	            			XPropertySet xPropSet = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xShape);
	            			String name = (String) xPropSet.getPropertyValue("LayerName");
	            			if(StringUtils.equals(myName, name))
	            			{
	            				Shape shape = null;
	            				
	            				//System.out.println("Layer: "+xShape.getShapeType());
	            				
	            				// Shape ist diesem Layer zugeordnet
	            				switch (xShape.getShapeType())
									{
										case Shape.LineShapeType:
											shape = new LineShape(xShape);											
											break;
											
										case Shape.PolyLineShapeType:
											shape = new PolyLineShape(xShape);															
											break;
											
										case Shape.OpenBezierType:
											shape = new OpenBezierShape(xShape);														
											break;
											
										case Shape.CustomType:
											shape = new Shape(xShape);														
											break;

										case Shape.TextShapeType:
											shape = new TextShape(xShape);														
											break;

										case Shape.EllipseShapeType:
											shape = new Shape(xShape);														
											break;

										case Shape.GraphicObjectShapeType:
											shape = new Shape(xShape);														
											break;

										case Shape.GroupShapeType:
											shape = new Shape(xShape);														
											break;


										default:
											break;
									}
	            				
	            				if(shape != null)
	            				{
	            					shape.layer = this;
	            					Point pos = xShape.getPosition();
	            					Size size =	xShape.getSize();
	            					shape.bound = new Rectangle(pos.X, pos.Y, size.Width, size.Height);
	            					shapes.add(shape);
	            				}
	            			}
	            		}
	            	}
	            }
	            
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}			
		}
		return shapes;
	}
	

}
