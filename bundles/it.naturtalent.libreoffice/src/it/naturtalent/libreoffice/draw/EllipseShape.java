package it.naturtalent.libreoffice.draw;

import it.naturtalent.libreoffice.Utils;

import org.apache.commons.lang3.ArrayUtils;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.graphics.Rectangle;

import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.uno.UnoRuntime;

public class EllipseShape extends Shape
{
	public EllipseShape(XShape xShape)
	{
		super(xShape);		
	}

	public EllipseShape(Layer layer, Rectangle bound) throws Exception
	{
		super(layer, bound, ShapeType.EllipseShape);		
	}

	public void setPolyPoints(Point polyPoints[])
	{
		if(ArrayUtils.isNotEmpty(polyPoints))
		{
			
		
			
			
			com.sun.star.awt.Point[][] awtPolyPoints = new com.sun.star.awt.Point[1][];
			com.sun.star.awt.Point[] awtPoints = null;  
			for(int n = 0;n < polyPoints.length;n++)
			{			
				com.sun.star.awt.Point awtPoint = new com.sun.star.awt.Point();
				awtPoint.X = polyPoints[n].x;
				awtPoint.Y = polyPoints[n].y;
				awtPoints = ArrayUtils.add(awtPoints, awtPoint);
			}
			awtPolyPoints[0] = awtPoints;
			
			try
			{
				XPropertySet xShapeProperties = UnoRuntime.queryInterface(XPropertySet.class, xShape);
				
				Utils.printPropertyValues(xShapeProperties);
				
				//xShapeProperties.setPropertyValue("PolyPolygon", awtPolyPoints);
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}


}
