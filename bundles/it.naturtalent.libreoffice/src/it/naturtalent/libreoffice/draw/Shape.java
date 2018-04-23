package it.naturtalent.libreoffice.draw;

import java.awt.geom.AffineTransform;

import it.naturtalent.libreoffice.ShapeHelper;

import org.eclipse.swt.graphics.Rectangle;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.HomogenMatrix3;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.view.XSelectionSupplier;

public class Shape implements IShape
{
	protected XShape xShape;
	protected Layer layer;	
	protected Rectangle bound;
	protected double translateX;	
	protected double translateY;	
	protected double rotate;
	private ShapeType shapeType;
	

	// xShape Defininitonen
	public static final String LineShapeType = "com.sun.star.drawing.LineShape";
	public static final String RectangleShapeType = "com.sun.star.drawing.RectangleShape";
	public static final String PolyLineShapeType = "com.sun.star.drawing.PolyLineShape";
	public static final String PolyPolygonType = "com.sun.star.drawing.PolyPolygonShape";
	public static final String OpenBezierType = "com.sun.star.drawing.OpenBezierShape";
	public static final String CustomType = "com.sun.star.drawing.CustomShape";
	public static final String TextShapeType = "com.sun.star.drawing.TextShape";
	public static final String EllipseShapeType = "com.sun.star.drawing.EllipseShape";
	public static final String GraphicObjectShapeType = "com.sun.star.drawing.GraphicObjectShape";
	public static final String GroupShapeType = "com.sun.star.drawing.GroupShape";
	
	public enum ShapeType
	{
		  LineShape, 
		  RectangleShape, 
		  PolyLineShape, 
		  PolyPolgonShape, 
		  OpenBezierShape, 
		  CustomShape,
		  EllipseShape,
		  GraphicObjectShape,
		  GroupShape,
		  TextShape;
		  

		  public String getType()
		  {
		    if ( this == LineShape)
		      return LineShapeType;
		    
		    if ( this == RectangleShape)
			      return RectangleShapeType;
		    
		    if ( this == PolyLineShape)
			      return PolyLineShapeType;

		    if ( this == PolyPolgonShape)
			      return PolyPolygonType;

		    if ( this == OpenBezierShape)
			      return OpenBezierType;
		    
		    if ( this == CustomShape)
			      return CustomType;

		    if ( this == TextShape)
			      return TextShapeType;
		    
		    if ( this == EllipseShape)
			      return EllipseShapeType;

		    if ( this == GraphicObjectShape)
			      return GraphicObjectShapeType;

		    if ( this == GroupShape)
			      return GroupShapeType;


		    return null;
		  }	
	}
	
	public Shape (XShape xShape)
	{
		super();
		this.xShape = xShape;
		
		//readShapeData();
	}
	
	/**
	 * Shape erzeugen und gleich der Seite hinzufuegen
	 * @param drawDocument
	 * @param bound
	 * @param shapeType
	 * @throws Exception
	 */
	public Shape (Layer layer, Rectangle bound, ShapeType shapeType ) throws Exception
    {
		this.layer = layer;
		this.bound = bound;		
		this.shapeType = shapeType;
		xShape = ShapeHelper.createShape(layer.getDrawDocument().getxComponent(),
				new Point(bound.x, bound.y),
				new Size(bound.width, bound.height), shapeType.getType());
				
		// Shape der Seite und dem Layer hinzufuegen
		layer.pushShape(this);
    }
	
	public String getShapeType()
	{
		return xShape.getShapeType();	
	}
	
	protected void transform()
	{
		try
		{
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
			HomogenMatrix3 aHomogenMatrix3 = (HomogenMatrix3)
		               xPropSet.getPropertyValue( "Transformation" );
			
			java.awt.geom.AffineTransform aOriginalMatrix =
	                   new java.awt.geom.AffineTransform(
	                       aHomogenMatrix3.Line1.Column1, aHomogenMatrix3.Line2.Column1,
	                       aHomogenMatrix3.Line1.Column2, aHomogenMatrix3.Line2.Column2,
	                       aHomogenMatrix3.Line1.Column3, aHomogenMatrix3.Line2.Column3 );

			 AffineTransform aNewMatrix1 = new AffineTransform();
	           aNewMatrix1.setToRotation(rotate);
	           aNewMatrix1.concatenate( aOriginalMatrix );

	           AffineTransform aNewMatrix2 = new AffineTransform();
	           aNewMatrix2.setToTranslation(translateX, translateY);
	           aNewMatrix2.concatenate( aNewMatrix1 );

	           double aFlatMatrix[] = new double[ 6 ];
	           aNewMatrix2.getMatrix( aFlatMatrix );
	           aHomogenMatrix3.Line1.Column1 = aFlatMatrix[ 0 ];
	           aHomogenMatrix3.Line2.Column1 = aFlatMatrix[ 1 ];
	           aHomogenMatrix3.Line1.Column2 = aFlatMatrix[ 2 ];
	           aHomogenMatrix3.Line2.Column2 = aFlatMatrix[ 3 ];
	           aHomogenMatrix3.Line1.Column3 = aFlatMatrix[ 4 ];
	           aHomogenMatrix3.Line2.Column3 = aFlatMatrix[ 5 ];
	           xPropSet.setPropertyValue( "Transformation", aHomogenMatrix3 );


		} catch (UnknownPropertyException | WrappedTargetException | IllegalArgumentException | PropertyVetoException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
	}
	
	public String getShapeLabel()
	{
		String label = "undefiniert";
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
	
	public void readShapeData()
	{
		try
		{
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
			String label =  (String) xPropSet.getPropertyValue("UINamePlural");
			
			com.sun.star.awt.Rectangle rect = (com.sun.star.awt.Rectangle) xPropSet.getPropertyValue("BoundRect");
			bound = new Rectangle(rect.X, rect.Y, rect.Width, rect.Height);
			
			//System.out.println(bound); 
			
			/*
			Property [] props = xPropSet.getPropertySetInfo().getProperties();
			for(Property prop : props)
			{
				System.out.println(prop.Name+" : "+xPropSet.getPropertyValue(prop.Name));
			}
			*/
			
		} catch (UnknownPropertyException | WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
			
	}

	
	
	public void select(DrawDocument drawDocument)
	{
		XModel xModel = UnoRuntime.queryInterface(XModel.class, drawDocument.xComponent);
		XController xController = xModel.getCurrentController();
		XSelectionSupplier xSelectionSupplier = UnoRuntime.queryInterface(
				XSelectionSupplier.class, xController);
		xSelectionSupplier.select(xShape);
	}

	public double getTranslateX()
	{
		return translateX;
	}

	public void setTranslateX(double translateX)
	{
		this.translateX = translateX;
	}

	public double getTranslateY()
	{
		return translateY;
	}

	public void setTranslateY(double translateY)
	{
		this.translateY = translateY;
	}

	public double getRotate()
	{
		return rotate;
	}

	public void setRotate(double rotate)
	{
		this.rotate = rotate;
	}

	public Rectangle getBound()
	{
		return bound;
	}

	

}
