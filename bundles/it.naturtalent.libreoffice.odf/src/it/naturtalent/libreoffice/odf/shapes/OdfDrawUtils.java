package it.naturtalent.libreoffice.odf.shapes;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.graphics.Rectangle;
import org.odftoolkit.odfdom.doc.OdfDocument;
import org.odftoolkit.odfdom.dom.OdfContentDom;
import org.odftoolkit.odfdom.dom.OdfStylesDom;
import org.odftoolkit.odfdom.dom.element.draw.DrawCustomShapeElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawFrameElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawLineElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawPathElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawPolylineElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawStrokeDashElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawTextBoxElement;
import org.odftoolkit.odfdom.dom.element.office.OfficeAutomaticStylesElement;
import org.odftoolkit.odfdom.dom.element.style.StyleGraphicPropertiesElement;
import org.odftoolkit.odfdom.dom.element.style.StyleStyleElement;
import org.odftoolkit.odfdom.dom.element.style.StyleTextPropertiesElement;
import org.odftoolkit.odfdom.dom.element.text.TextPElement;
import org.odftoolkit.odfdom.dom.element.text.TextSpanElement;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeAutomaticStyles;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.odftoolkit.odfdom.pkg.OdfElement;
import org.odftoolkit.odfdom.pkg.OdfFileDom;
import org.odftoolkit.simple.Document;
import org.w3c.dom.Node;

public class OdfDrawUtils
{
	
	public static final String ELLIPSE_SHAPE_DEFINITIONTYPE = "ellipse";	
	public static final String RECTANGLE_SHAPE_DEFINITIONTYPE = "rectangle";
	
	

	// FillAttribute
	public static final String FILL_NONE = "none";
	public static final String FILL_SOLID = "solid";
	public static final String FILL_BITMAP = "bitmap";
	public static final String FILL_GRADIENT = "gradient";
	
	// Linestyle
	public static final String LINE_NONE = "none";
	public static final String LINE_SOLID = "solid";
	public static final String LINE_DASH = "dash";

	static Map<String,Double>unitFactors = new HashMap<String, Double>();
	static
	{
		unitFactors.put("cm", 1000.0);
	}

	/**
	 * @param drawFrameElement
	 * @return
	 */
	public static Rectangle getLineShapeBBx(DrawLineElement drawLineElement)
	{
		String x1 = drawLineElement.getSvgX1Attribute();
		x1 = StringUtils.isEmpty(x1) ? "0.0" : x1;
				
		String y1 = drawLineElement.getSvgY1Attribute();
		y1 = StringUtils.isEmpty(y1) ? "0.0" : y1;

		String x2 = drawLineElement.getSvgX2Attribute();
		x2 = StringUtils.isEmpty(x2) ? "0.0" : x2;
				
		String y2 = drawLineElement.getSvgY2Attribute();
		y2 = StringUtils.isEmpty(y2) ? "0.0" : y2;
			
		ODFRectangle odfBound = new ODFRectangle(x1,y1,x2,y2);
		return getBound(odfBound);
	}

	/**
	 * @param drawFrameElement
	 * @return
	 */
	public static Rectangle getCustormShapeBBx(DrawCustomShapeElement drawCustomShapeElement)
	{
		String x = drawCustomShapeElement.getSvgXAttribute();
		x = StringUtils.isEmpty(x) ? "0.0" : x;
				
		String y = drawCustomShapeElement.getSvgYAttribute();
		y = StringUtils.isEmpty(y) ? "0.0" : y;
		
		String width = drawCustomShapeElement.getSvgWidthAttribute();
		String height = drawCustomShapeElement.getSvgHeightAttribute();
			
		ODFRectangle odfBound = new ODFRectangle(x,y,width,height);
		return getBound(odfBound);
	}

	/**
	 * @param drawFrameElement
	 * @return
	 */
	public static Rectangle getDrawPolyLinesBBx(DrawPolylineElement drawPolylineElement)
	{
		String x = drawPolylineElement.getSvgXAttribute();
		x = StringUtils.isEmpty(x) ? "0.0" : x;
				
		String y = drawPolylineElement.getSvgYAttribute();
		y = StringUtils.isEmpty(y) ? "0.0" : y;
		
		String width = drawPolylineElement.getSvgWidthAttribute();
		String height = drawPolylineElement.getSvgHeightAttribute();
			
		ODFRectangle odfBound = new ODFRectangle(x,y,width,height);
		return getBound(odfBound);
	}
	
	/**
	 * @param drawFrameElement
	 * @return
	 */
	public static Rectangle getDrawPathBBx(DrawPathElement drawPathElement)
	{
		String x = drawPathElement.getSvgXAttribute();
		x = StringUtils.isEmpty(x) ? "0.0" : x;
				
		String y = drawPathElement.getSvgYAttribute();
		y = StringUtils.isEmpty(y) ? "0.0" : y;
		
		String width = drawPathElement.getSvgWidthAttribute();
		String height = drawPathElement.getSvgHeightAttribute();
			
		ODFRectangle odfBound = new ODFRectangle(x,y,width,height);
		return getBound(odfBound);
	}


	/**
	 * @param drawFrameElement
	 * @return
	 */
	public static Rectangle getDrawFrameBBx(DrawFrameElement drawFrameElement)
	{
		String x = drawFrameElement.getSvgXAttribute();
		x = StringUtils.isEmpty(x) ? "0.0" : x;
				
		String y = drawFrameElement.getSvgYAttribute();
		y = StringUtils.isEmpty(y) ? "0.0" : y;
		
		String width = drawFrameElement.getSvgWidthAttribute();
		String height = drawFrameElement.getSvgHeightAttribute();
			
		ODFRectangle odfBound = new ODFRectangle(x,y,width,height);
		return getBound(odfBound);
	}
	
	/**
	 * @param odfBound
	 * @return
	 */
	public static Rectangle getBound(ODFRectangle odfBound)
	{
		// BoundedBox
		int x = convertToInteger(odfBound.x);
		int y = convertToInteger(odfBound.y);
		int width = convertToInteger(odfBound.width);
		int height = convertToInteger(odfBound.height);
		return new Rectangle(x, y, width, height);
	}
	
	

	/**
	 * @param measurement
	 * @return
	 */
	public static Integer convertToInteger(ODFMeasurement measurement)
	{
		Double dblVal = measurement.value;
		Double factor = unitFactors.get(measurement.measureUnit);
		if(factor != null)
			dblVal = dblVal * factor;
		return dblVal.intValue();
	}
	
	/**
	 * @param stgPolyPoints
	 * @return
	 */
	public static  Point [] convertToPoints(String stgPolyPoints)
	{
		Point [] points = null;
		
		if(StringUtils.isNotEmpty(stgPolyPoints))
		{
			String [] pointArray = StringUtils.split(stgPolyPoints, " ");			
			for(int i = 0;i < pointArray.length;i++)
			{				
				String [] xyArray = StringUtils.split(pointArray[i], ",");
				
				int x = new Integer(xyArray[0]).intValue();
				int y = new Integer(xyArray[1]).intValue();
				Point point = new Point(x,y);
				points = ArrayUtils.add(points, point);
			}
		}
		
		return points;
	}


	
	/**
	 * Rueckgabe der zu 'drawFrameElement' gehoerenden Styles
	 * @param drawFrameElement
	 * @return
	 */
	public static StyleTextPropertiesElement getStyleTextPropertiesElement(
			DrawFrameElement drawFrameElement)
	{
		// den TextStyleName ermitteln
		DrawTextBoxElement textBoxElement = OdfElement.findFirstChildNode(
				DrawTextBoxElement.class, drawFrameElement);
		
		if (textBoxElement == null)
			return null;

		TextPElement textPElement = OdfElement.findFirstChildNode(
				TextPElement.class, textBoxElement);

		TextSpanElement textSpanElement = OdfElement.findFirstChildNode(
				TextSpanElement.class, textPElement);

		String textStyleName = textSpanElement.getTextStyleNameAttribute();

		OfficeAutomaticStylesElement stylesElement = drawFrameElement
				.getAutomaticStyles();
		StyleStyleElement styleStyleElement = OdfStyle.findFirstChildNode(
				StyleStyleElement.class, stylesElement);
		
		// das entsprechende StyleElement ermitteln
		StyleStyleElement siblingElement = styleStyleElement;
		do
		{				
			siblingElement = OdfElement.findNextChildNode(StyleStyleElement.class, siblingElement);
			if(StringUtils.equals(siblingElement.getStyleNameAttribute(), textStyleName))
				break;
			
		} while (siblingElement != null);
		
		if(siblingElement != null)
			return OdfStyle.findFirstChildNode(StyleTextPropertiesElement.class, siblingElement);
			
		return null;
	} 
	
	/**
	 * Sammelt eine Auswahl der definierten Styles auf und gibt sie in der Datenstruktur 
	 * DrawOdfStyle zurueck;
	 * 
	 * @param odfStyle
	 * @return
	 */
	public static DrawOdfStyle getDrawOdfStyles(OdfOfficeStyles odfOfficeStyle, OdfStyle odfStyle)
	{
				
		DrawOdfStyle drawOdfStyle = new DrawOdfStyle();
		StyleGraphicPropertiesElement styleGraphicPropertyElement = OdfStyle.findFirstChildNode(
				StyleGraphicPropertiesElement.class, odfStyle);
		
		// Fuellstyle
		drawOdfStyle.fillAttribute = styleGraphicPropertyElement.getDrawFillAttribute();
		drawOdfStyle.fillAttribute = StringUtils
				.isEmpty(drawOdfStyle.fillAttribute) ? FILL_SOLID
				: drawOdfStyle.fillAttribute;
		
		// Fuellfarbe
		String stgFillColor = styleGraphicPropertyElement.getDrawFillColorAttribute();
		if (StringUtils.isNotEmpty(stgFillColor))
			drawOdfStyle.fillColor = Integer.decode(stgFillColor);

		// Linienbreite
		ODFMeasurement lineWidthMeasure = new ODFMeasurement();
		String stgLineWidth = styleGraphicPropertyElement.getSvgStrokeWidthAttribute();
		if(StringUtils.isNotEmpty(stgLineWidth))
		{
			lineWidthMeasure.parseInput(styleGraphicPropertyElement
					.getSvgStrokeWidthAttribute());
			drawOdfStyle.lineWidth = convertToInteger(lineWidthMeasure);
		}
		
		// Linienstil
		drawOdfStyle.lineStyle = styleGraphicPropertyElement.getDrawStrokeAttribute();
		drawOdfStyle.lineStyle = (StringUtils.isNotEmpty(drawOdfStyle.lineStyle)) ? drawOdfStyle.lineStyle : LINE_SOLID;
		
		// Liniefarbe
		String stgLineColor = styleGraphicPropertyElement.getSvgStrokeColorAttribute();
		if (StringUtils.isNotEmpty(stgLineColor))
			drawOdfStyle.lineColor = Integer.decode(stgLineColor);

		// gestrichelte Linie			
		if(StringUtils.equals(drawOdfStyle.lineStyle, LINE_DASH))
		{
			drawOdfStyle.lineDash = true;
			String dashName = styleGraphicPropertyElement.getDrawStrokeDashAttribute();
						
			// existiert eine StrokeDashElement
			DrawStrokeDashElement strokeDashElement = findStrokeDashElement(odfOfficeStyle, dashName);
			if(strokeDashElement != null)
			{
				// StrokeDashDaten einlesen
				drawOdfStyle.strokeDashDefinition = parseStrokeDashElement(strokeDashElement, drawOdfStyle.lineWidth);
			}
		}
		
		return drawOdfStyle;
	}
	
	public static DrawOdfStyle addTransformStyles(DrawOdfStyle drawOdfStyle, String transform)
	{
		if(StringUtils.isNotEmpty(transform))
		{
			// Rotate			
			String stgRotate = StringUtils.substringAfter(transform, "rotate");
			if (StringUtils.isNotEmpty(stgRotate))
			{
				stgRotate = StringUtils.remove(stgRotate, ")");
				stgRotate = StringUtils.remove(stgRotate, "(");
				String[] arrayRotate = StringUtils.split(stgRotate);
				drawOdfStyle.rotate = new Double(arrayRotate[0]);
			}

			// Translate
			String stgTranslate = StringUtils.substringAfter(transform, "translate");
			if(StringUtils.isNotEmpty(stgTranslate))
			{
				String[] splitTranslate = StringUtils.split(stgTranslate);			
				splitTranslate[1] = StringUtils.remove(splitTranslate[1], ")");

				// x
				ODFMeasurement translateX = new ODFMeasurement();
				translateX.parseInput(splitTranslate[0]);
				Double dblTranslateX = translateX.value;
				Double factor = unitFactors.get(translateX.measureUnit);
				dblTranslateX = dblTranslateX * factor;
				drawOdfStyle.translate = ArrayUtils.add(drawOdfStyle.translate, dblTranslateX);

				// y
				ODFMeasurement translateY = new ODFMeasurement();
				translateY.parseInput(splitTranslate[1]);
				Double dblTranslateY = translateY.value;				
				factor = unitFactors.get(translateY.measureUnit);			
				dblTranslateY = dblTranslateY * factor;
				drawOdfStyle.translate = ArrayUtils.add(drawOdfStyle.translate, dblTranslateY);
			}
		}
		
		return drawOdfStyle;
	}
	
	/*
	 * Rueckgabe des unter dem Namen 'dashName' eingetragenen DrawStrokeDashElements. 
	 * 'odfOfficeStyle' Styleknoten in 'styles.xml'.
	 */
	private static DrawStrokeDashElement findStrokeDashElement(OdfOfficeStyles odfOfficeStyle, String dashName)
	{
		DrawStrokeDashElement strokeDashElement = OdfElement
				.findFirstChildNode(DrawStrokeDashElement.class,odfOfficeStyle);
		
		if(strokeDashElement != null)
		{
			if (!StringUtils.equals(strokeDashElement.getDrawNameAttribute(),dashName))
			{
				do
				{
					strokeDashElement = OdfElement.findNextChildNode(DrawStrokeDashElement.class, strokeDashElement);
					if(strokeDashElement != null)
					{
						if(StringUtils.equals(strokeDashElement.getDrawNameAttribute(),dashName))
							break;
					}
				} while (strokeDashElement != null);
			}
		}
		
		return strokeDashElement;
	}
	
	/*
	 * Die StrokeDashinformationen in eine eigene Struktur uebernehmen und zurueckgeben.
	 */
	private static StrokeDashDefinition parseStrokeDashElement(DrawStrokeDashElement strokeDashElement, Integer lineWidth)
	{
		StrokeDashDefinition strokeDashDefinition = new StrokeDashDefinition();
		strokeDashDefinition.dots1 = strokeDashElement.getDrawDots1Attribute();
		strokeDashDefinition.dots2 = strokeDashElement.getDrawDots2Attribute();

		ODFMeasurement measure = new ODFMeasurement();
		String dotLen = strokeDashElement.getDrawDots1LengthAttribute();
		if(StringUtils.isNotEmpty(dotLen))
		{
			measure.parseInput(dotLen);			
			if(!measure.measureUnit.equals("%"))			
				strokeDashDefinition.dotlen = convertToInteger(measure);
			else
			{
				int width = (lineWidth == null) ? 270 : lineWidth.intValue();
				strokeDashDefinition.dotlen = measure.value.intValue() * width / 100;
			}
		}
		
		String distance = strokeDashElement.getDrawDistanceAttribute();
		if(StringUtils.isNotEmpty(distance))
		{
			measure.parseInput(distance);
			if(!measure.measureUnit.equals("%"))			
				strokeDashDefinition.distance = convertToInteger(measure);
			else
			{
				int width = (lineWidth == null) ? 270 : lineWidth.intValue();
				strokeDashDefinition.distance = measure.value.intValue() * width / 100;
			}
		}
		
		return strokeDashDefinition;
	}
	
}
