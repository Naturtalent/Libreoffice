package it.naturtalent.libreoffice.draw;

import it.naturtalent.libreoffice.Utils;
import it.naturtalent.libreoffice.math.BigDecimalMath;

import java.math.BigDecimal;

import org.apache.commons.lang3.ArrayUtils;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.graphics.Rectangle;

import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.PolyPolygonBezierCoords;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;

public class OpenBezierShape extends Shape
{
	public enum PolygonFlags
	{
		  NORMAL, SMOOTH, CONTROL, SYMMETRIC;

		  public com.sun.star.drawing.PolygonFlags getValue()
		  {
		    if ( this == NORMAL)
		    	return  com.sun.star.drawing.PolygonFlags.NORMAL;
		    
		    if ( this == SMOOTH)
		    	return com.sun.star.drawing.PolygonFlags.SMOOTH;
		    
		    if ( this == CONTROL)
			      return com.sun.star.drawing.PolygonFlags.CONTROL;

		    if ( this == SYMMETRIC)
			      return com.sun.star.drawing.PolygonFlags.SYMMETRIC;

		    return null;
		  }	
	}

	
	public OpenBezierShape(XShape xShape)
	{
		super(xShape);		
	}
	
	public OpenBezierShape(Layer layer, Rectangle bound) throws Exception
	{
		super(layer, bound, ShapeType.OpenBezierShape);
		// TODO Auto-generated constructor stub
	}
	
	public void setPolyPoints(Point coordinates[][], PolygonFlags flags[][])
	{
		 PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();
		 
		 com.sun.star.awt.Point[][] coords = new com.sun.star.awt.Point[coordinates.length][];
		 for(int n = 0; n < coordinates.length;n++)
		 {
			 Point [] coordPoints = coordinates[n];
			 com.sun.star.awt.Point[] pPolyPoints = new com.sun.star.awt.Point[coordPoints.length];
			 for(int i=0;i < pPolyPoints.length; i++)
				 pPolyPoints[i] = new com.sun.star.awt.Point();
		
			 for(int i = 0;i < coordPoints.length;i++)
			 {
				 pPolyPoints[i].X = coordPoints[i].x;
				 pPolyPoints[i].Y = coordPoints[i].y;
			 }			 
			 coords[n] = pPolyPoints;
		 }		 
		 aCoords.Coordinates = coords;

		 com.sun.star.drawing.PolygonFlags [][] polygonFlags = new com.sun.star.drawing.PolygonFlags[flags.length][];
		 for(int n = 0; n < polygonFlags.length;n++)
		 {
			 PolygonFlags [] polyFlags = flags[n];
			 com.sun.star.drawing.PolygonFlags [] pPolyFlag = new com.sun.star.drawing.PolygonFlags[polyFlags.length];
			 
			 for(int i = 0;i < pPolyFlag.length;i++)				 
				 pPolyFlag[i] = polyFlags[i].getValue();
			 
			 polygonFlags[n] = pPolyFlag;
		 }
		 
		aCoords.Coordinates = coords;
		aCoords.Flags = polygonFlags;

		try
		{
			XPropertySet xShapeProperties = UnoRuntime.queryInterface(
					XPropertySet.class, xShape);
			xShapeProperties.setPropertyValue("PolyPolygonBezier", aCoords);

		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
	 * Konvertiert die Bezierkurve in PolyLines
	 * 
	 * @return
	 */
	public Point [] getPolyLines()
	{		
		Point[] bezierPoints = getBezierPoints();
		if (ArrayUtils.isNotEmpty(bezierPoints))
		{
			BezierCurve bc = new BezierCurve();
			bc.convertToPoly(bezierPoints);
			return bc.getPolyLines();
		}
		return null;
	}

	/**
	 * Laenge der Bezierkurve ermitteln. 
	 * 
	 * @return
	 */
	public BigDecimal getLength()
	{
		BigDecimal sum = null;
		try
		{
			// temporaeres PolyLineShape generieren
			PolyLineShape polyLineShape = new PolyLineShape(layer, bound);
			
			// mit den konvertierten Bezierpolylines fuellen
			Point [] polyLines = getPolyLines();
			polyLineShape.setPolyPoints(polyLines);
			
			// Laenge ermitteln
			sum = polyLineShape.getLength();
			
			// temp. Polylineshape wieder loeschen
			layer.deleteShape(polyLineShape);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return sum;		
	}

	public Point[] getBezierPoints()
	{
		try
		{
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);		
			PolyPolygonBezierCoords awtCoord = (PolyPolygonBezierCoords) xPropSet.getPropertyValue("PolyPolygonBezier");
			com.sun.star.awt.Point [][] awtPoints = awtCoord.Coordinates;
				
			if (ArrayUtils.isNotEmpty(awtPoints))
			{
				Point[] coord = null;
				for (int n = 0; n < awtPoints[0].length; n++)
				{
					Point point = new Point(awtPoints[0][n].X,
							awtPoints[0][n].Y);
					coord = ArrayUtils.add(coord, point);
				}				
				return coord;
			}
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	public Point[][] getBezierPointsOLD()
	{
		try
		{
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);		
			PolyPolygonBezierCoords awtCoord = (PolyPolygonBezierCoords) xPropSet.getPropertyValue("PolyPolygonBezier");
			com.sun.star.awt.Point [][] awtPoints = awtCoord.Coordinates;
				
			if (ArrayUtils.isNotEmpty(awtPoints))
			{
				Point[][] coords = new Point[1][];
				Point[] coord = null;
				for (int n = 0; n < awtPoints[0].length; n++)
				{
					Point point = new Point(awtPoints[0][n].X,
							awtPoints[0][n].Y);
					coord = ArrayUtils.add(coord, point);
				}
				coords[0] = coord;
				return coords;
			}
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}
	
	public BigDecimal getLengthOLD()
	{		
		BigDecimal sum = null;
		try
		{
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
			
			PolyPolygonBezierCoords coord = (PolyPolygonBezierCoords) xPropSet.getPropertyValue("PolyPolygonBezier");
			com.sun.star.awt.Point [][] points = coord.Coordinates;
			com.sun.star.drawing.PolygonFlags [][] flags = coord.Flags;
			
			for(int i = 0;i < points.length;i++)
			{
				System.out.println("Polygon Points");
				com.sun.star.awt.Point [] pPolyPoints = points[i];				
				for(int n = 0;n < pPolyPoints.length;n++)
					System.out.println("Point: "+n+"  "+pPolyPoints[n].X+" :  "+pPolyPoints[n].Y);
			}

			for(int i = 0;i < flags.length;i++)
			{
				System.out.println("Polygon Flags");
				com.sun.star.drawing.PolygonFlags [] pPolyFlags = flags[i];				
				for(int n = 0;n < pPolyFlags.length;n++)
				{
					String stgFlag = "";
					switch (pPolyFlags[n].getValue())
						{
							case com.sun.star.drawing.PolygonFlags.CONTROL_value:
								stgFlag = "Control";
								break;

							case com.sun.star.drawing.PolygonFlags.NORMAL_value:
								stgFlag = "Normal";
								break;
								
							case com.sun.star.drawing.PolygonFlags.SMOOTH_value:
								stgFlag = "Smooth";
								break;

							case com.sun.star.drawing.PolygonFlags.SYMMETRIC_value:
								stgFlag = "Symmetric";
								break;

							default:
								break;
						}
					
					
					System.out.println("Flag: "+n+"  "+stgFlag);
				}
			}

			
			//PolyPolygonBezierCoords coord2 = (PolyPolygonBezierCoords) xPropSet.getPropertyValue("Geometry");
			//System.out.println("COORD2: "+coord2.Coordinates.length);
			
			/*
			Point[] points = (Point[]) xPropSet.getPropertyValue("Polygon");
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
					len = BigDecimalMath.sqrt(len); 

					if(sum == null) 
						sum = len;
					else
						sum = sum.add(len);
				}
				return sum;
			}
			*/
			
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
				
		return sum;
	}
	

}
