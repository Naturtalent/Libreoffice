package it.naturtalent.libreoffice.draw;

import java.util.HashMap;
import java.util.Map;

import it.naturtalent.libreoffice.Utils;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.LineDash;
import com.sun.star.drawing.LineStyle;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;

public class Style
{

	// Styleobjectnames
	public static final String STYLEOBJECT_STANDARD = "standard";
	public static final String STYLEOBJECT_WITHOUTFILL = "objectwithoutfill";
	
	private String styleObject = STYLEOBJECT_STANDARD;
	
	// Linedash definitionen
	public static final String LINESTYLE_NONE = "linestylenone";
	public static final String LINESTYLE_SOLID = "linestylesolid";
	public static final String LINESTYLE_DASH = "linestyledash";
	

	private Map<String,LineStyle>lineStyleMap = new HashMap<String, LineStyle>();
	{
		lineStyleMap.put(LINESTYLE_NONE, LineStyle.NONE);
		lineStyleMap.put(LINESTYLE_SOLID, LineStyle.SOLID);
		lineStyleMap.put(LINESTYLE_DASH, LineStyle.DASH);
	}
	
	private IItemStyle itemStyle;
	
	
	public void setItemStyle(IItemStyle itemStyle)
	{
		this.itemStyle = itemStyle;
	}
	
	/**
	 * Die definierten Eigenschaften an das Dokument uebertragen-
	 * 
	 * @param drawDocument
	 */
	public void pushItemStyle(DrawDocument drawDocument)
	{
		if(itemStyle != null)
		{
			try
			{
				setStyleObject(Style.STYLEOBJECT_WITHOUTFILL);
				XPropertySet xStylePropSet = getStylePropSet(drawDocument);
				
				// Linienfarbe
				Integer lineColor = itemStyle.getLineColor();	
				if(lineColor != null)					
					xStylePropSet.setPropertyValue("LineColor", lineColor);
				
				// Linienstarke
				Integer lineWidth = itemStyle.getLineWidth();
				if(lineWidth != null)
					xStylePropSet.setPropertyValue("LineWidth", lineWidth);
				
				// Linienstiel
				xStylePropSet.setPropertyValue("LineStyle",com.sun.star.drawing.LineStyle.SOLID);

				LineDash aLineDash = new LineDash();	
				if (itemStyle.getLinedash() != null)
				{
					aLineDash.Dashes = itemStyle.getLinedash().getDashes();
					aLineDash.DashLen = itemStyle.getLinedash().getDashlen();
					aLineDash.Distance = itemStyle.getLinedash().getDistance();
					aLineDash.Dots = itemStyle.getLinedash().getDots();
					aLineDash.DotLen = itemStyle.getLinedash().getDotlen();

					xStylePropSet.setPropertyValue("LineStyle",
							com.sun.star.drawing.LineStyle.DASH);
					xStylePropSet.setPropertyValue("LineDash", aLineDash);
				}
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}


	//
	// Linestyle
	//
	private LineStyle lineStyle;
	
	public LineStyle getLineStyle(String styleType)
	{
		return lineStyle;
	}

	public void setLineStyle(String lineStype)
	{
		this.lineStyle = lineStyle;
	}
	
	public void pullLineStyle(DrawDocument drawDocument) 
	{		
		try
		{
			XPropertySet xStylePropSet = getStylePropSet(drawDocument);
			lineStyle = (LineStyle) xStylePropSet.getPropertyValue("LineStyle");			
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}	
		
	public void pushLineStyle(DrawDocument drawDocument) 
	{		
		try
		{
			XPropertySet xStylePropSet = getStylePropSet(drawDocument);
			xStylePropSet.setPropertyValue("LineStyle", lineStyle);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}	


	//
	// Linecolor
	//
	private Integer lineColor;
	
	public Integer getLineColor()
	{
		return lineColor;
	}

	public void setLineColor(Integer lineColor)
	{
		this.lineColor = lineColor;
	}
	
	public void pullLineColor(DrawDocument drawDocument) 
	{		
		try
		{
			XPropertySet xStylePropSet = getStylePropSet(drawDocument);
			lineColor = (Integer) xStylePropSet.getPropertyValue("LineColor");			
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}	
		
	public void pushLineColor(DrawDocument drawDocument) 
	{		
		try
		{
			XPropertySet xStylePropSet = getStylePropSet(drawDocument);
			xStylePropSet.setPropertyValue("LineColor", lineColor);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}	

	//
	// Linewidth
	//
	private Integer lineWidth;
	
	public Integer getLineWidth()
	{
		return lineWidth;
	}

	public void setLineWidth(Integer lineWidth)
	{
		this.lineWidth = lineWidth;
	}
	
	public void pullLineWidth(DrawDocument drawDocument) 
	{		
		try
		{
			XPropertySet xStylePropSet = getStylePropSet(drawDocument);
			lineWidth = (Integer) xStylePropSet.getPropertyValue("LineWidth");
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}	
		
	public void pushLineWidth(DrawDocument drawDocument) 
	{		
		try
		{
			XPropertySet xStylePropSet = getStylePropSet(drawDocument);
			xStylePropSet.setPropertyValue("LineWidth", lineWidth);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}	

	//
	// LineDash
	//

	
	
	
	
	//
	//
	//
	public void printProperties(DrawDocument drawDocument)
	{
		XPropertySet xPropertySet = getStylePropSet(drawDocument);
		Utils.printPropertyValues(xPropertySet);
	}
	
	
	private XPropertySet getStylePropSet(DrawDocument drawDocument)
	{
		try
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class, drawDocument.xComponent);
			com.sun.star.style.XStyleFamiliesSupplier xSFS = UnoRuntime
					.queryInterface(
							com.sun.star.style.XStyleFamiliesSupplier.class, xModel);
			com.sun.star.container.XNameAccess xFamilies = xSFS.getStyleFamilies();

			Object obj = xFamilies.getByName("graphics");
			com.sun.star.container.XNameAccess xStyles = UnoRuntime.queryInterface(
					com.sun.star.container.XNameAccess.class, obj);
			
			String [] names = xStyles.getElementNames();

			Object aStyleObj = xStyles.getByName(styleObject);
			//Object aStyleObj = xStyles.getByName("standard");
			//Object aStyleObj = xStyles.getByName("objectwithoutfill");
			
			com.sun.star.style.XStyle xStyle = UnoRuntime.queryInterface(
					com.sun.star.style.XStyle.class, aStyleObj);
			
			return UnoRuntime.queryInterface(XPropertySet.class, xStyle);

		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	public String getStyleObject()
	{
		return styleObject;
	}

	public void setStyleObject(String styleObject)
	{
		this.styleObject = styleObject;
	}
	
	

}
