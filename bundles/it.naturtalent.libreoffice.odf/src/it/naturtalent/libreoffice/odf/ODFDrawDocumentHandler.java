package it.naturtalent.libreoffice.odf;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.eclipse.swt.graphics.Point;
import org.odftoolkit.odfdom.dom.element.draw.DrawCustomShapeElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawFrameElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawGElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawLineElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawPageElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawPathElement;
import org.odftoolkit.odfdom.dom.element.draw.DrawPolylineElement;
import org.odftoolkit.odfdom.dom.element.office.OfficeDrawingElement;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.pkg.OdfElement;
import org.odftoolkit.simple.GraphicsDocument;

public class ODFDrawDocumentHandler extends ODFOfficeDocumentHandler
{
	private GraphicsDocument graphicsDocument;
	
	// Zugriff auf den Inhalt 'content.xml'
	private OfficeDrawingElement content = null;
	
	// Zugriff auf die Styles 'styles.xml'
	private OdfOfficeStyles styles = null;
	
	@Override
	public void openOfficeDocument(File fileDoc)
	{
		try
		{
			if ((fileDoc != null) && (fileDoc.exists()))
			{
				if(graphicsDocument != null)
					graphicsDocument.close();
					
				graphicsDocument = null;
				content = null;
									
				graphicsDocument = GraphicsDocument.loadDocument(fileDoc);
				if(graphicsDocument != null)
				{
					content = graphicsDocument.getContentRoot();					
					styles = graphicsDocument.getDocumentStyles();
				}
			}
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	@Override
	public void saveOfficeDocument()
	{
		if (graphicsDocument != null)
		{
			try
			{
				graphicsDocument.save(fileDoc);
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	@Override
	public void closeOfficeDocument()
	{
		if (graphicsDocument != null)
		{
			graphicsDocument.close();
			graphicsDocument = null;
		}
	}
	
	@Override
	public Object getDocument()
	{		
		return graphicsDocument;
	}
	
	/**
	 * Die Seiten eines DrawDocuments in ein Array einlesen.
	 * 
	 * @return
	 */
	public DrawPageElement [] getDrawPages()
	{
		DrawPageElement [] pages = null;
				
		DrawPageElement pageElement = OdfElement.findFirstChildNode(DrawPageElement.class, content);
		if(pageElement != null)
		{
			pages = ArrayUtils.add(pages, pageElement);
			
			DrawPageElement siblingPageElement = pageElement;
			do
			{
				siblingPageElement = OdfElement.findNextChildNode(DrawPageElement.class, siblingPageElement);
				if(siblingPageElement != null)
					pages = ArrayUtils.add(pages, siblingPageElement);
				
			} while (siblingPageElement != null);
		}
		
		return pages;
	}
	
	/**
	 * Eine bestimmte Seite des DrawDocuments zurueckgeben
	 * 
	 * @return
	 */
	public DrawPageElement getDrawPage(String pageName)
	{
		DrawPageElement [] pages = getDrawPages();
		if(ArrayUtils.isNotEmpty(pages))
		{
			if(StringUtils.isEmpty(pageName))
				return pages[0];
						
			for(DrawPageElement page : pages)
			{
				String name = page.getDrawNameAttribute();
				if(StringUtils.equals(name, pageName))
					return page;
			}				
		}
		
		return null;
	}
	
	/**
	 * Alle Elemente mit Layerattribute *layerName' der Seite 'pageName' auflisten.
	 * 
	 * @param pageName
	 * @return
	 */
	public List<OdfElement> getPageLayerElements(String pageName, String layerName)
	{
		List<OdfElement>odfElements = null;
				
		DrawPageElement drawPageElement = getDrawPage(pageName);
		if(drawPageElement != null)
		{
			// die aufnehmende Liste
			odfElements = new ArrayList<OdfElement>();
			
			// rekursiver Durchlauf
			listLayerChildElement(drawPageElement, odfElements, layerName);				
		}
		
		return odfElements;
	}
	
	/*
	 * Tree unterhalb DrawPage durchlaufen und Layerelemente aufspueren
	 */
	private void listLayerChildElement(OdfElement parentElement, List<OdfElement>lElements,String layerName)
	{
		int size = 0;
		
		OdfElement childElement = OdfElement.findFirstChildNode(OdfElement.class, parentElement);
		if(childElement != null)
		{
			if(addLayer(childElement, lElements, layerName) == null)
			{
				// Group-start
				if(childElement instanceof DrawGElement)
				{
					lElements.add(childElement);
					size = lElements.size();
				}
				
				// weiter eine Ebene tiefer
				listLayerChildElement(childElement, lElements, layerName);
				
				// Group-end
				if(childElement instanceof DrawGElement)
				{
					if(lElements.size() == size)
					{
						// empty Group eliminieren
						size--;
						lElements.remove(size);
					}
					else lElements.add(childElement);
				}
			}
								
			OdfElement siblingElement = childElement;			
			do
			{				
				siblingElement = OdfElement.findNextChildNode(OdfElement.class, siblingElement);
				if(siblingElement != null)
				{				
					if(addLayer(siblingElement, lElements, layerName) == null)
					{
						// Group-start
						if(siblingElement instanceof DrawGElement)
						{
							lElements.add(siblingElement);
							size = lElements.size();
						}
						
						listLayerChildElement(siblingElement, lElements, layerName);
						
						// Group-end
						if(siblingElement instanceof DrawGElement)
						{
							if(lElements.size() == size)
							{
								// empty Group eliminieren
								size--;
								lElements.remove(size);
							}
							else lElements.add(siblingElement);
						}
					}
				}				
			} while (siblingElement != null);
		}
	}
	
	/*
	 * zur Liste hinzufuegen, wenn definiertes Element erkannt wird
	 */
	private OdfElement addLayer(OdfElement odfElement, List<OdfElement>lElements, String layerName)
	{
		String checkName;
		
		if (odfElement instanceof DrawCustomShapeElement)						
		{
			DrawCustomShapeElement shapeElement = (DrawCustomShapeElement) odfElement;
			checkName = shapeElement.getDrawLayerAttribute();
			if(StringUtils.equals(checkName, layerName))
			{
				lElements.add(shapeElement);
				return odfElement;
			}			
		}

		if (odfElement instanceof DrawPathElement)						
		{
			DrawPathElement shapeElement = (DrawPathElement) odfElement;
			checkName = shapeElement.getDrawLayerAttribute();
			if(StringUtils.equals(checkName, layerName))
			{
				lElements.add(shapeElement);
				return odfElement;
			}
		}
		
		if (odfElement instanceof DrawLineElement)						
		{
			DrawLineElement shapeElement = (DrawLineElement) odfElement;
			checkName = shapeElement.getDrawLayerAttribute();
			if(StringUtils.equals(checkName, layerName))
			{
				lElements.add(shapeElement);
				return odfElement;
			}	
		}
		
		if (odfElement instanceof DrawPolylineElement)						
		{
			DrawPolylineElement shapeElement = (DrawPolylineElement) odfElement;
			checkName = shapeElement.getDrawLayerAttribute();
			if(StringUtils.equals(checkName, layerName))
			{
				lElements.add(shapeElement);
				return odfElement;
			}
		}

		if (odfElement instanceof DrawFrameElement)						
		{
			DrawFrameElement shapeElement = (DrawFrameElement) odfElement;
			checkName = shapeElement.getDrawLayerAttribute();
			if(StringUtils.equals(checkName, layerName))
			{
				lElements.add(shapeElement);
				return odfElement;
			}
		}


		return null;
	}

	/**
	 * Alle Layer der Seite 'pageName' auflisten
	 * 
	 * @return
	 */
	public DrawCustomShapeElement [] getPageLayers(String pageName)
	{
		DrawCustomShapeElement [] layers = null;
				
		DrawPageElement drawPageElement = getDrawPage(pageName);
		if(drawPageElement != null)
		{		
			DrawCustomShapeElement layer = OdfElement.findFirstChildNode(DrawCustomShapeElement.class, drawPageElement);
			if(layer != null)
			{
				layers = ArrayUtils.add(layers, layer);
				
				DrawCustomShapeElement siblingLayer = layer;
				do
				{
					siblingLayer = OdfElement.findNextChildNode(DrawCustomShapeElement.class, siblingLayer);
					if(siblingLayer != null)
						layers = ArrayUtils.add(layers, siblingLayer);
					
				} while (siblingLayer != null);
			}
		}

		return layers;
	}

	
	/**
	 * Alle Layer der Seite 'pageName' mit dem Namen 'layerName' auflisten
	 * 
	 * @return
	 */
	/*
	public DrawCustomShapeElement [] getPageLayers(String pageName, String layerName)
	{
		DrawCustomShapeElement [] pageLayers = getPageLayers(pageName);
		DrawCustomShapeElement [] layers = null;
		
		if(ArrayUtils.isNotEmpty(pageLayers))
		{
			for(DrawCustomShapeElement pageLayer : pageLayers)
			{
				String name = pageLayer.getDrawLayerAttribute();
				if(StringUtils.equals(name, layerName))
					layers = ArrayUtils.add(layers, pageLayer);
			}
		}
			
		
		return layers;
	}
	*/

	

	public OdfOfficeStyles getStyles()
	{
		return styles;
	}

	
	private Point [] convertToPoints(String stgPolyPoints)
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

}
