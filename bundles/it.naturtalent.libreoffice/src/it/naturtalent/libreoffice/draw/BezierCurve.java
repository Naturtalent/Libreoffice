package it.naturtalent.libreoffice.draw;

import java.lang.Math;
import java.util.ArrayList;
import java.util.List;

import org.eclipse.swt.graphics.Point;

public class BezierCurve
{

	private double m_distance_tolerance = 0.25;
	
	private List<Point>points = new ArrayList<Point>();

	public void convertToPoly(Point[] bezierPoints)
	{		
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
	}

	public void convertToPolyOLD(Point[][] bezierPoints)
	{
		int diffNdx = 4;
		for(int i = 0;i < bezierPoints.length;i+=diffNdx)
		{
			double x1 = bezierPoints[0][i].x;
			double y1 = bezierPoints[0][i].y;
			double x2 = bezierPoints[0][i+1].x;
			double y2 = bezierPoints[0][i+1].y;
			double x3 = bezierPoints[0][i+2].x;
			double y3 = bezierPoints[0][i+2].y;
			double x4 = bezierPoints[0][i+3].x;
			double y4 = bezierPoints[0][i+3].y;
			
			add_point(x1, y1);
			recursive_bezier(x1, y1, x2, y2, x3, y3, x4, y4);
			add_point(x4, y4);
		}
	}
	
	void recursive_bezier(double x1, double y1, double x2, double y2,
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
	
	public Point[] getPolyLines()
	{
		/*
		Point[][] coords = new Point[1][];
		Point [] polyPoints = points.toArray(new Point[points.size()]);
		coords[0] = polyPoints;
		*/
		
		return points.toArray(new Point[points.size()]);
	}
	
}
