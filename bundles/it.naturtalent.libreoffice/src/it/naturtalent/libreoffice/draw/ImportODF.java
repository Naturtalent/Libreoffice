package it.naturtalent.libreoffice.draw;

import it.naturtalent.libreoffice.draw.Shape.ShapeType;
import it.naturtalent.libreoffice.odf.ODFDrawDocumentHandler;
import it.naturtalent.libreoffice.odf.shapes.DrawOdfStyle;
import it.naturtalent.libreoffice.odf.shapes.OdfDrawUtils;
import it.naturtalent.libreoffice.odf.shapes.StrokeDashDefinition;

import java.awt.geom.AffineTransform;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.graphics.Rectangle;
import org.odftoolkit.odfdom.dom.element.draw.DrawCustomShapeElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawEnhancedGeometryElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawFrameElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawGElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawLineElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawPathElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawPolylineElement;
import org.odftoolkit.odfdom.dom.element.style.StyleTextPropertiesElement;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.odftoolkit.odfdom.pkg.OdfElement;

import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.HomogenMatrix3;
import com.sun.star.drawing.LineDash;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.uno.UnoRuntime;

public class ImportODF
{
	
	//private ODFDrawDocumentHandler odfDocumentHandler  = new ODFDrawDocumentHandler();
	
	private ODFDrawDocumentHandler odfDocumentHandler;
	
	private List<OdfElement>odfElements;
	
	private Shape groupShape;
	
	/**
	 * Quelldatei mit ODFDrawDocumentHandler oeffnen.
	 * @param sourceDrawDocumentPath
	 */
	/*
	public void openDrawDocument(String sourceDrawDocumentPath)
	{		
		File file = new File(sourceDrawDocumentPath);
		if(file.exists())				
			odfDocumentHandler.openOfficeDocument(file);
	}
	*/

	
	/**
	 * Die ODFElemente der Ebene (Layers) eine ausgewaehlten Seite zusammenstellen.
	 *  
	 * @param pageName
	 * @param layerName
	 */
	/*
	public void readLayerShapes(String pageName, String layerName)
	{		
		odfElements = odfDocumentHandler.getPageLayerElements(pageName, layerName);
		if(odfElements != null)
		{			
			for(OdfElement element : odfElements)
				System.out.println(element);			
		}		
	}
	*/
	
	
	/**
	 * Die mit 'readLayerShapes()' zusammengestellten ODFElemente in Ziellayer 'layer' importieren.
	 * @param layer
	 */
	public void importShapes(Layer layer)
	{
		if((layer != null) && (odfElements != null) && (!odfElements.isEmpty()))
		{
			importOdfElements(layer, odfElements.iterator(), null);
		}
	}

	private void importOdfElements(Layer layer, Iterator<OdfElement>itOdfElement, OdfElement groupElement)
	{
		Shape shape = null;	
		Shape curGroupShape = null;

		if(groupElement != null)
		{
			try
			{
				curGroupShape = new Shape(layer,
						new Rectangle(0, 0, 0, 0), ShapeType.GroupShape);
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
				
		while (itOdfElement.hasNext())
		{
			OdfElement odfElement = itOdfElement.next();

			// GroupElement
			if (odfElement instanceof DrawGElement)
			{
				if((groupElement != null) && odfElement.equals(groupElement))
				{
					// Ende einer Gruppe erreicht
					return;
				}
				
				// Start einer neuen Gruppe
				importOdfElements(layer, itOdfElement, odfElement);
			}

			shape = importShape(layer, odfElement);

			// zur Gruppe hinzufuegen
			if ((curGroupShape != null) && (shape != null))
			{
				XShape xGroup = curGroupShape.xShape;
				XShapes xShapesGroup = UnoRuntime.queryInterface(XShapes.class,xGroup);
				xShapesGroup.add(shape.xShape);
			}

		}

	}	
	
	private Shape importShape(Layer layer, OdfElement odfElement)
	{
		Shape shape = null;	
		
		// DrawLineShapeElement
		if (odfElement instanceof DrawLineElement)
			shape = createLineShapeElement(layer,
					(DrawLineElement) odfElement);

		// DrawCustomShapeElement
		if (odfElement instanceof DrawCustomShapeElement)
			shape = createCustomShapeElement(layer,
					(DrawCustomShapeElement) odfElement);

		// DrawPolyLineElement
		if (odfElement instanceof DrawPolylineElement)
			shape = createPolyLineShape(layer,
					(DrawPolylineElement) odfElement);

		// DrawFrameElement
		if (odfElement instanceof DrawFrameElement)
			shape = createFrameShape(layer, (DrawFrameElement) odfElement);

		// DrawPathElement
		if (odfElement instanceof DrawPathElement)
			createPathShape(layer, (DrawPathElement) odfElement);

		return shape;
	}

	private Shape createLineShapeElement(Layer layer, DrawLineElement drawLineElement)
	{
		LineShape shape = null;
		
		try
		{
			Rectangle bbx = OdfDrawUtils.getLineShapeBBx(drawLineElement);			
			shape = new LineShape(layer, bbx);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		
		//transform(drawLineShapeDefinition, shape);
		// Styles uebernehmen
		OdfStyle odfStyle = drawLineElement.getAutomaticStyle();			
		OdfOfficeStyles odfOfficeStyle = odfDocumentHandler.getStyles();
		DrawOdfStyle drawOdfStyle = OdfDrawUtils.getDrawOdfStyles(odfOfficeStyle, odfStyle);
		style(drawOdfStyle, shape);
		
		return shape;
	}
	
	// PolyLineShape erzeugen
	private Shape createPolyLineShape(Layer layer, DrawPolylineElement drawPolylineElement)
	{
		PolyLineShape shape = null;
		try
		{	
			// PolyLineShape erzeugen
			Rectangle bbx = OdfDrawUtils.getDrawPolyLinesBBx(drawPolylineElement);			
			shape = new PolyLineShape(layer, bbx);
			
			// Polygonpunkte setzen
			String stgPolyPoints = drawPolylineElement.getDrawPointsAttribute();
			Point [] polyPoints = OdfDrawUtils.convertToPoints(stgPolyPoints);
			shape.setPolyPoints(polyPoints);
			
			// Styles uebernehmen
			OdfStyle odfStyle = drawPolylineElement.getAutomaticStyle();			
			OdfOfficeStyles odfOfficeStyle = odfDocumentHandler.getStyles();
			DrawOdfStyle drawOdfStyle = OdfDrawUtils.getDrawOdfStyles(odfOfficeStyle, odfStyle);
			style(drawOdfStyle, shape);
			
			// Transform
			String transform = drawPolylineElement.getDrawTransformAttribute();
			if (StringUtils.isNotEmpty(transform))
			{
				OdfDrawUtils.addTransformStyles(drawOdfStyle, transform);
				
				// ToDo Transform realisieren
			}
			else
			{
				// PolyLine in BBx verschieben
				AffineTransform affineTransform = getTransformMatrix(shape);
				affineTransform.setToTranslation(bbx.x, bbx.y);
				setTransformMatrix(affineTransform, shape);
			}
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return shape;
	}
	
	// PathShape erzeugen
	private Shape createPathShape(Layer layer, DrawPathElement drawPathElement)
	{
		OpenBezierShape shape = null;
		try
		{	
			// OpenBezierShape erzeugen
			Rectangle bbx = OdfDrawUtils.getDrawPathBBx(drawPathElement);			
			shape = new OpenBezierShape(layer, bbx);
			
			// Polygonpunkte setzen
			String stgPolyPoints = drawPathElement.getSvgDAttribute();
			
			//Point [] polyPoints = OdfDrawUtils.convertToPoints(stgPolyPoints);
			//shape.setPolyPoints(polyPoints);
			System.out.println(stgPolyPoints);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return shape;
	}


	private Shape createCustomShapeElement(Layer layer, DrawCustomShapeElement drawCustomShapeElement)
	{
		Shape shape = null;
		
		try
		{		
			DrawEnhancedGeometryElement drawEnhancedGeometry = OdfElement
					.findFirstChildNode(DrawEnhancedGeometryElement.class,
							drawCustomShapeElement);

			Rectangle bbx = OdfDrawUtils.getCustormShapeBBx(drawCustomShapeElement);
			switch (drawEnhancedGeometry.getDrawTypeAttribute())
				{					
					case OdfDrawUtils.ELLIPSE_SHAPE_DEFINITIONTYPE:
						shape = new EllipseShape(layer, bbx);								
						break;

					case OdfDrawUtils.RECTANGLE_SHAPE_DEFINITIONTYPE:
						shape = new RectangleShape(layer,bbx);
						break;

					default:
						break;
				}
			
			// Styles uebernehmen
			OdfStyle odfStyle = drawCustomShapeElement.getAutomaticStyle();			
			OdfOfficeStyles odfOfficeStyle = odfDocumentHandler.getStyles();
			DrawOdfStyle drawOdfStyle = OdfDrawUtils.getDrawOdfStyles(odfOfficeStyle, odfStyle);
			style(drawOdfStyle, shape);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return shape;
	}
	
	// TextFrame
	private Shape createFrameShape(Layer layer, DrawFrameElement drawFrameElement)
	{
		TextShape shape = null;
		try
		{
			Rectangle bbx = OdfDrawUtils.getDrawFrameBBx(drawFrameElement);
			shape = new TextShape(layer, bbx);			
			shape.setText(drawFrameElement.getTextContent());
			styleText(drawFrameElement, shape);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
				
		return shape;
	}

	
	private AffineTransform getTransformMatrix(Shape shape)
	{
		XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, shape.xShape);
		if(xPropertySet != null)
		{
			try
			{
				HomogenMatrix3 aHomogenMatrix3 = (HomogenMatrix3)
						xPropertySet.getPropertyValue( "Transformation" );
				
				AffineTransform aOriginalMatrix =
				           new java.awt.geom.AffineTransform(
				               aHomogenMatrix3.Line1.Column1, aHomogenMatrix3.Line2.Column1,
				               aHomogenMatrix3.Line1.Column2, aHomogenMatrix3.Line2.Column2,
				               aHomogenMatrix3.Line1.Column3, aHomogenMatrix3.Line2.Column3 );				
				return aOriginalMatrix;
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		return null;
	}
	
	private void setTransformMatrix(AffineTransform affineTransform, Shape shape)
	{
		XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, shape.xShape);
		if(xPropertySet != null)
		{
	        try
			{
	        	HomogenMatrix3 aHomogenMatrix3 = (HomogenMatrix3)
						xPropertySet.getPropertyValue( "Transformation" );
	        	
				double aFlatMatrix[] = new double[ 6 ];
				affineTransform.getMatrix( aFlatMatrix );
				aHomogenMatrix3.Line1.Column1 = aFlatMatrix[ 0 ];
				aHomogenMatrix3.Line2.Column1 = aFlatMatrix[ 1 ];
				aHomogenMatrix3.Line1.Column2 = aFlatMatrix[ 2 ];
				aHomogenMatrix3.Line2.Column2 = aFlatMatrix[ 3 ];
				aHomogenMatrix3.Line1.Column3 = aFlatMatrix[ 4 ];
				aHomogenMatrix3.Line2.Column3 = aFlatMatrix[ 5 ];
				xPropertySet.setPropertyValue( "Transformation", aHomogenMatrix3 );
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}	
	}
	
	/*
	private void transform(Shape shape)
	{
		XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, shape.xShape);
		if(xPropertySet != null)
		{
			Double [] translate = drawShapeDefinition.getTranslate();
			if(ArrayUtils.isNotEmpty(translate))
			{
				try
				{
					
					HomogenMatrix3 aHomogenMatrix3 = (HomogenMatrix3)
							xPropertySet.getPropertyValue( "Transformation" );
		
					
					java.awt.geom.AffineTransform aOriginalMatrix =
				               new java.awt.geom.AffineTransform(
				                   aHomogenMatrix3.Line1.Column1, aHomogenMatrix3.Line2.Column1,
				                   aHomogenMatrix3.Line1.Column2, aHomogenMatrix3.Line2.Column2,
				                   aHomogenMatrix3.Line1.Column3, aHomogenMatrix3.Line2.Column3 );
					
					AffineTransform rotateMatrix = null;
					Double rotate = drawShapeDefinition.getRotate();
					
					if(rotate != null)
					{
						rotate = rotate * (-1.0);
						rotateMatrix = new AffineTransform();
						rotateMatrix.setToRotation(rotate);
						rotateMatrix.concatenate( aOriginalMatrix );
					}
					
					AffineTransform translateMatrix = new AffineTransform();
			        translateMatrix.setToTranslation(translate[0], translate[1]);
			        
			        if(rotateMatrix != null)
			        	translateMatrix.concatenate(rotateMatrix);
			        else			        
			        	translateMatrix.concatenate(aOriginalMatrix);
			         
			        double aFlatMatrix[] = new double[ 6 ];
			        translateMatrix.getMatrix( aFlatMatrix );
			        aHomogenMatrix3.Line1.Column1 = aFlatMatrix[ 0 ];
			        aHomogenMatrix3.Line2.Column1 = aFlatMatrix[ 1 ];
			        aHomogenMatrix3.Line1.Column2 = aFlatMatrix[ 2 ];
			        aHomogenMatrix3.Line2.Column2 = aFlatMatrix[ 3 ];
			        aHomogenMatrix3.Line1.Column3 = aFlatMatrix[ 4 ];
			        aHomogenMatrix3.Line2.Column3 = aFlatMatrix[ 5 ];
			        xPropertySet.setPropertyValue( "Transformation", aHomogenMatrix3 );
					
				} catch (Exception e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}
	*/

	private void styleText(DrawFrameElement drawFrameElement, Shape shape)
	{
		// Style Definitionen
		if (shape != null)
		{
			try
			{
				XPropertySet xPropertySet = UnoRuntime.queryInterface(
						XPropertySet.class, shape.xShape);
				if (xPropertySet != null)
				{					
					StyleTextPropertiesElement styleTextProperty = OdfDrawUtils.getStyleTextPropertiesElement(drawFrameElement);
					
					if (styleTextProperty != null)
					{
						// TextWeigt
						String textWeight = styleTextProperty
								.getFoFontWeightAttribute();
						switch (textWeight)
							{
								case "bold":
									xPropertySet.setPropertyValue("CharWeight",new Float(com.sun.star.awt.FontWeight.BOLD));
									break;

								default:
									break;
							}
					}
				}
			}catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	private void style(DrawOdfStyle drawOdfStyle, Shape shape)
	{
		if (shape != null)
		{
			try
			{
				XPropertySet xPropertySet = UnoRuntime.queryInterface(
						XPropertySet.class, shape.xShape);
				if (xPropertySet != null)
				{
					// FillStyle definieren
					FillStyle fillstyle = FillStyle.NONE;
					String fillAttribute = drawOdfStyle.fillAttribute;
					if (StringUtils.isNotEmpty(fillAttribute))
					{
						switch (fillAttribute)
							{
								case OdfDrawUtils.FILL_SOLID:
									fillstyle = FillStyle.SOLID;
									break;

								case OdfDrawUtils.FILL_BITMAP:
									fillstyle = FillStyle.BITMAP;
									break;

								case OdfDrawUtils.FILL_GRADIENT:
									fillstyle = FillStyle.GRADIENT;
									break;

								default:
									break;
							}
					}
					xPropertySet.setPropertyValue("FillStyle", fillstyle);

					// FillColor
					Integer fillColor = drawOdfStyle.fillColor;
					if (fillColor != null)
						xPropertySet.setPropertyValue("FillColor", fillColor);

					// LineWidth
					Integer lineWidth = drawOdfStyle.lineWidth;
					if (lineWidth != null)
						xPropertySet.setPropertyValue("LineWidth", lineWidth);

					// LineColor
					Integer lineColor = drawOdfStyle.lineColor;
					if (lineColor != null)
						xPropertySet.setPropertyValue("LineColor", lineColor);

					// LineStyle definieren
					LineStyle lineStyle = LineStyle.SOLID;
					String styleAttribute = drawOdfStyle.lineStyle;
					if (StringUtils.isNotEmpty(styleAttribute))
					{
						switch (styleAttribute)
							{
								case OdfDrawUtils.LINE_DASH:
									lineStyle = LineStyle.DASH;
									break;

								case OdfDrawUtils.LINE_NONE:
									lineStyle = LineStyle.NONE;
									break;
							}
					}
					xPropertySet.setPropertyValue("LineStyle", lineStyle);

					// LineDash
					StrokeDashDefinition strokeShapeDefinition = drawOdfStyle.strokeDashDefinition;
					if (strokeShapeDefinition != null)
					{
						LineDash aLineDash = new LineDash();
						aLineDash.Dots = strokeShapeDefinition.dots1
								.shortValue();
						aLineDash.Dashes = strokeShapeDefinition.dots2
								.shortValue();
						aLineDash.DotLen = strokeShapeDefinition.dotlen
								.intValue();
						aLineDash.Distance = strokeShapeDefinition.distance;
						// aLineDash.DashLen =
						// itemStyle.getLinedash().getDotlen();

						xPropertySet.setPropertyValue("LineDash", aLineDash);
					}

				}
			}catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	public void setOdfElements(List<OdfElement> odfElements)
	{
		this.odfElements = odfElements;
	}

	public void setOdfDocumentHandler(ODFDrawDocumentHandler odfDocumentHandler)
	{
		this.odfDocumentHandler = odfDocumentHandler;
	}

	/**
	 * Die mit 'readLayerShapes()' zusammengestellten ODFElemente in Ziellayer 'layer' importieren.
	 * @param layer
	 */
	public void importShapesOLD(Layer layer)
	{
		Shape shape = null;
		
		if((layer != null) && (odfElements != null) && (!odfElements.isEmpty()))
		{
			for(OdfElement odfElement : odfElements)
			{
				// GroupElement
				if (odfElement instanceof DrawGElement)
				{
					try
					{
						if(groupShape == null)						
							groupShape = new Shape(layer,new Rectangle(0,0,0,0),ShapeType.GroupShape);						
						else 
							groupShape = null;
						
					} catch (Exception e)
					{
						// TODO Auto-generated catch block
						e.printStackTrace();
					}					
				}
	
				// DrawLineShapeElement
				if (odfElement instanceof DrawLineElement)
					shape = createLineShapeElement(layer, (DrawLineElement)odfElement);
	
				// DrawCustomShapeElement
				if (odfElement instanceof DrawCustomShapeElement)
					shape = createCustomShapeElement(layer, (DrawCustomShapeElement)odfElement);
	
				// DrawPolyLineElement
				if (odfElement instanceof DrawPolylineElement)				
					shape = createPolyLineShape(layer, (DrawPolylineElement)odfElement);
				
				// DrawFrameElement
				if (odfElement instanceof DrawFrameElement)
					shape = createFrameShape(layer, (DrawFrameElement)odfElement);
	
				// DrawPathElement				
				if (odfElement instanceof DrawPathElement)
					createPathShape(layer, (DrawPathElement)odfElement);
					
				// zur Gruppe hinzufuegen					
				if ((groupShape != null) && (shape != null))
				{
					XShape xGroup = groupShape.xShape;
					XShapes xShapesGroup = UnoRuntime.queryInterface(
							XShapes.class, xGroup);
					xShapesGroup.add(shape.xShape);
				}
			}
		}
	}

	

}
