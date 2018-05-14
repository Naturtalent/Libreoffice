package it.naturtalent.libreoffice;

/**
 * Definiert eine unspezifische Shape-Klasse. 
 * 
 * Mit der Funktion @see DrawDocument.getLayerShapes(String pageName, String layerName)
 * werden die spezifischen xShapes gelesen. Fuer jedes xShape wird ein DrawShape generiert
 * mit dem Object xShape als Inhalt.
 * 
 * @author dieter
 *
 */
public class DrawShape implements IDrawShape
{
	
	// LibreOffice - spezifische xShape-Types
	/*
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
	*/

	/* 
	 * Die definierten unspezifische ShapeTypes
	 * 
	 * 
	 */
	public enum SHAPEPROP
	{
		Label,		// DrawShape repraesentiert nur den Label
		Line		// DrawShape repraesentiert die Länge einer Linie-/Linienzuges
	}
	
	// LibreOffice - spezifische Klasse
	public Object xShape;     
		
	// Defaulttype
	private SHAPEPROP drawShapeType = SHAPEPROP.Label;
	
	// Bound
	private int x;
	private int y;
	private int widht;
	private int height;

	// interner Datenspeicher
	private Object data;
		
	/**
	 * Konstruktion
	 * 
	 * @param x
	 * @param y
	 * @param widht
	 * @param height
	 */
	public DrawShape(int x, int y, int widht, int height)
	{
		super();
		this.x = x;
		this.y = y;
		this.widht = widht;
		this.height = height;
	}
	
	public DrawShape()
	{
		super();
	}



	public SHAPEPROP getDrawShapeType()
	{
		return drawShapeType;
	}

	public void setDrawShapeType(SHAPEPROP drawShapeType)
	{
		this.drawShapeType = drawShapeType;
	}
	
	

	public int getX()
	{
		return x;
	}

	public void setX(int x)
	{
		this.x = x;
	}

	public int getY()
	{
		return y;
	}

	public void setY(int y)
	{
		this.y = y;
	}

	public int getWidht()
	{
		return widht;
	}

	public void setWidht(int widht)
	{
		this.widht = widht;
	}
	
	public int getHeight()
	{
		return height;
	}

	public void setHeight(int height)
	{
		this.height = height;
	}
	
	public Object getxShape()
	{
		return xShape;
	}

	public void setxShape(Object xShape)
	{
		this.xShape = xShape;
	}

	public Object getData()
	{
		return data;
	}

	public void setData(Object data)
	{
		this.data = data;
	}

	
	 
	
	
	
}
