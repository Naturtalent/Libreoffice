package it.naturtalent.libreoffice.draw;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.eclipse.swt.graphics.Point;

import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.HomogenMatrix3;
import com.sun.star.drawing.PolyPolygonBezierCoords;
import com.sun.star.drawing.XLayer;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;

import it.naturtalent.libreoffice.DrawDocumentUtils;
import it.naturtalent.libreoffice.DrawShape;
import it.naturtalent.libreoffice.math.BigDecimalMath;

public class ShapeLineLengthUtils
{
	
	/*
	public ShapeLineLengthUtils(DrawDocument drawDocument)
	{
		super();
		this.drawDocument = drawDocument;
	}
	*/

	public BigDecimal getLineLength(XShape xShape)
	{
		String type = xShape.getShapeType();
		if(StringUtils.equals(type, DrawDocumentUtils.LineShapeType))
		{
			return getLength(xShape);
			//return formatLength (length);			
		}
		
		if(StringUtils.equals(type, DrawDocumentUtils.PolyLineShapeType))
		{
			return getPolyLength(xShape);
		}

		if(StringUtils.equals(type, DrawDocumentUtils.OpenBezierType))
		{
			return getBezierLength(xShape);
		}

		
		return new BigDecimal(0.0);
	}
	
	/*
	 * 
	 *  einfache Linie
	 * 
	 * 
	 * 
	 */
	
	/**
	 * Laenge einer Linie
	 * @param xShape
	 * @return
	 */
	public BigDecimal getLength(XShape xShape)
	{
		XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(
				XPropertySet.class, xShape);
		try
		{
			HomogenMatrix3 aHomogenMatrix3 = (HomogenMatrix3) xPropSet
					.getPropertyValue("Transformation");

			BigDecimal breit = new BigDecimal(aHomogenMatrix3.Line1.Column1);
			BigDecimal  hoch = new BigDecimal(aHomogenMatrix3.Line2.Column2);	
			
			breit = BigDecimalMath.powRound(breit,2);
			hoch = BigDecimalMath.powRound(hoch,2);
			return BigDecimalMath.sqrt(breit.add(hoch ));
			
		} catch (UnknownPropertyException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return new BigDecimal(0.0);
	}
	
	/*
	 * 
	 *  Polylinienzug
	 * 
	 * 
	 * 
	 */
	
	/**
	 * Laenge eines Linienzuges
	 * @param xShape
	 * @return
	 */
	public BigDecimal getPolyLength(XShape xShape)
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
				
		return new BigDecimal(0.0);
	}
	
	/*
	 * 
	 *  Bezierkurve
	 * 
	 * 
	 * 
	 */

	/**
	 * Laenge der Bezierkurve ermitteln. 
	 * 
	 * @return
	 */
	public BigDecimal getBezierLength(XShape xShape)
	{
		BigDecimal sum = null;
		try
		{
			// Bezier in PolyLine konvertieren  
			Point [] polyPoints = convertToPolyLines(xShape);

			// PolyLinePunke in awt Punkte konvertieren
			com.sun.star.awt.Point[] awtPoints = null;  
			for(int n = 0;n < polyPoints.length;n++)
			{			
				com.sun.star.awt.Point awtPoint = new com.sun.star.awt.Point();
				awtPoint.X = polyPoints[n].x;
				awtPoint.Y = polyPoints[n].y;
				awtPoints = ArrayUtils.add(awtPoints, awtPoint);
			}
			
			// Laenge des konvertierten PolyLines ermitteln 
			if (ArrayUtils.isNotEmpty(awtPoints))
			{
				int n = awtPoints.length - 1;
				for (int i = 0; i < n; i++)
				{
					BigDecimal lenX = new BigDecimal(awtPoints[i + 1].X - awtPoints[i].X);
					lenX = BigDecimalMath.powRound(lenX,2);

					BigDecimal lenY = new BigDecimal(awtPoints[i + 1].Y - awtPoints[i].Y);
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
		
		return new BigDecimal(0.0);
	}
	
	/**
	 * Konvertiert die Bezierkurve in PolyLines.
	 * 
	 * @return
	 */
	private double m_distance_tolerance = 0.25;
	private List<Point>points = new ArrayList<Point>();
	private Point [] convertToPolyLines(XShape xShape)
	{		
		// die Bezierpunkte einlesen
		Point[] bezierPoints = getBezierPoints(xShape);
		
		int n = bezierPoints.length -3;
		for(int i = 0;i < n;i+=3)
		{
			double x1 = bezierPoints[i].x;
			double y1 = bezierPoints[i].y;
			double x2 = bezierPoints[i+1].x;
			double y2 = bezierPoints[i+1].y;
			double x3 = bezierPoints[i+2].x;
			double y3 = bezierPoints[i+2].y;
			double x4 = bezierPoints[i+3].x;
			double y4 = bezierPoints[i+3].y;
			
			add_point(x1, y1);
			recursive_bezier(x1, y1, x2, y2, x3, y3, x4, y4);
			add_point(x4, y4);
		}
		
		return points.toArray(new Point[points.size()]);
	}
	
	private void recursive_bezier(double x1, double y1, double x2, double y2,
			double x3, double y3, double x4, double y4)
	{
	       // Calculate all the mid-points of the line segments
        //----------------------
        double x12   = (x1 + x2) / 2;
        double y12   = (y1 + y2) / 2;
        double x23   = (x2 + x3) / 2;
        double y23   = (y2 + y3) / 2;
        double x34   = (x3 + x4) / 2;
        double y34   = (y3 + y4) / 2;
        double x123  = (x12 + x23) / 2;
        double y123  = (y12 + y23) / 2;
        double x234  = (x23 + x34) / 2;
        double y234  = (y23 + y34) / 2;
        double x1234 = (x123 + x234) / 2;
        double y1234 = (y123 + y234) / 2;

        // Try to approximate the full cubic curve by a single straight line
        //------------------
        double dx = x4-x1;
        double dy = y4-y1;

        double d2 = Math.abs(((x2 - x4) * dy - (y2 - y4) * dx));
        double d3 = Math.abs(((x3 - x4) * dy - (y3 - y4) * dx));

        if((d2 + d3)*(d2 + d3) < m_distance_tolerance * (dx*dx + dy*dy))
        {
            add_point(x1234, y1234);
            return;
        }

        // Continue subdivision
        //----------------------
        recursive_bezier(x1, y1, x12, y12, x123, y123, x1234, y1234); 
        recursive_bezier(x1234, y1234, x234, y234, x34, y34, x4, y4);     
	}	
	
	private void add_point(double x, double y)
	{
		Point point = new Point((int)x, (int)y);
		points.add(point);
	}

	// die Bezierpunkte ermitteln
	private Point[] getBezierPoints(XShape xShape)
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



}
