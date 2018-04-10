package it.naturtalent.libreoffice.draw;

import it.naturtalent.libreoffice.math.BigDecimalMath;

import java.math.BigDecimal;

import org.apache.commons.lang3.ArrayUtils;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.graphics.Rectangle;

import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.uno.UnoRuntime;

public class PolyLineShape extends Shape
{

	public PolyLineShape(XShape xShape)
	{
		super(xShape);		
	}
	
	public PolyLineShape(Layer layer, Rectangle bound) throws Exception
	{
		super(layer, bound, ShapeType.PolyLineShape);	
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
				xShapeProperties.setPropertyValue("PolyPolygon", awtPolyPoints);
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	
	public BigDecimal getLength()
	{		
		BigDecimal sum = null;
		try
		{
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
			com.sun.star.awt.Point[] points = (com.sun.star.awt.Point[]) xPropSet.getPropertyValue("Polygon");
			if (ArrayUtils.isNotEmpty(points))
			{
				int n = points.length - 1;
				for (int i = 0; i < n; i++)
				{
					BigDecimal lenX = new BigDecimal(points[i + 1].X - points[i].X);
					lenX = BigDecimalMath.powRound(lenX,2);

					BigDecimal lenY = new BigDecimal(points[i + 1].Y - points[i].Y);
					lenY = BigDecimalMath.powRound(lenY,2);

					BigDecimal len = lenX.add(lenY);
					try
					{
						len = BigDecimalMath.sqrt(len);
						if(sum == null) 
							sum = len;
						else
							sum = sum.add(len);
					} catch (Exception e)
					{
					} 
				}
				return sum;
			}
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
				
		return sum;
	}
	

}
