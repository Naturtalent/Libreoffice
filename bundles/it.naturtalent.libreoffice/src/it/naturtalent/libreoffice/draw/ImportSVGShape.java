package it.naturtalent.libreoffice.draw;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.graphics.Rectangle;

import it.naturtalent.libreoffice.Utils;
import it.naturtalent.libreoffice.draw.Shape.ShapeType;

import com.sun.star.beans.PropertyValue;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.io.XInputStream;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.xml.sax.InputSource;
import com.sun.star.xml.sax.Parser;
import com.sun.star.xml.sax.SAXException;
import com.sun.star.xml.sax.XAttributeList;
import com.sun.star.xml.sax.XDocumentHandler;
import com.sun.star.xml.sax.XLocator;
import com.sun.star.xml.sax.XParser;

public class ImportSVGShape implements XDocumentHandler
{

	//String path = "/media/dieter/f8ceb1a1-74b6-4dbf-a487-e12e6249ced01/home/dieter/temp/DrawSvg1.svg";

	private XComponentContext xComponentContext;
	private String importPath;
	private Layer layer;
	
	private String inUseElement = null;
	private ShapeType detectedShapeType;
	private XAttributeList attributeList;
	
	
	
	public ImportSVGShape(Layer layer, String importPath)
	{
		super();
		this.layer = layer;
		this.importPath = importPath;
		xComponentContext = Utils.getxContext();		
	}
	
	public void readShapes()
	{
		if ((xComponentContext != null) && StringUtils.isNotEmpty(importPath))
		{
			XParser parser = Parser
					.create((XComponentContext) xComponentContext);

			try
			{
				File file = new File(importPath);
				InputStream is = new FileInputStream(file);
				XInputStream xStream = new ByteArrayToXInputStreamAdapter(
						IOUtils.toByteArray(is));

				InputSource aInput = new InputSource();
				aInput.sSystemId = "InputStream";
				aInput.aInputStream = xStream;
				parser.setDocumentHandler(this);
				parser.parseStream(aInput);
			} catch (IOException | Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	@Override
	public void characters(String arg0) throws SAXException
	{
		// TODO Auto-generated method stub
		
	}


	@Override
	public void endDocument() throws SAXException
	{
		//System.out.println("end");
		
	}


	@Override
	public void startElement(String arg0, XAttributeList arg1)
			throws SAXException
	{
	
		if(detectedShapeType != null)
		{
			if(attributeList == null)
			{
				if(StringUtils.equals(arg0,"path"))
					attributeList = arg1;			
			}			
		}
		else
		{
			ShapeType shapeType = detectShapeType(arg0, arg1);
			if (shapeType != null)
			{
				detectedShapeType = shapeType;
				inUseElement = arg0;
			}
		}
		
	}


	@Override
	public void endElement(String arg0) throws SAXException
	{
		if(StringUtils.equals(inUseElement, arg0))
		{
			if(detectedShapeType != null)
			{
				generateShape(detectedShapeType, attributeList);
				
				/*
				System.out.println(detectedShapeType);				
				short n = attributeList.getLength();
				for(short i = 0;i < n;i++)
				{
					String attribute = attributeList.getNameByIndex(i);
					String value = attributeList.getValueByIndex(i);
					System.out.println("Attribute: "+attribute+"  value: "+value);
				}
				*/
			}
			
			inUseElement = null;
			detectedShapeType = null;
			attributeList = null;
		}
	}

	private void generateShape(ShapeType shapeType, XAttributeList attributeList)
	{
		switch (shapeType)
			{
				case LineShape:
					generateLineShape(attributeList);
					break;
					
				case EllipseShape:
					generateEllipseShape(attributeList);
					break;


				default:
					break;
			}
	}

	private void generateEllipseShape(XAttributeList attributeList)
	{		
		try
		{
			Rectangle bound = new Rectangle(0, 0, 0, 0);
			EllipseShape ellipseShape = new EllipseShape(layer, bound);
			
			Point [] points = parse_d_Tag(attributeList);	
			ellipseShape.setPolyPoints(points);
			
			System.out.println(points);
			
			
		} catch (java.lang.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void generateLineShape(XAttributeList attributeList)
	{		
		try
		{
			Rectangle bound = new Rectangle(0, 0, 0, 0);
			LineShape lineShape = new LineShape(layer, bound);
			
			
			
			
		} catch (java.lang.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private Point [] parse_d_Tag(XAttributeList attributeList)
	{
		List<Point>points = new ArrayList<Point>();
		
		String stgAttb = attributeList.getValueByName("d");		
		String [] items = StringUtils.split(stgAttb," ");
		for(String item : items)
		{
			String [] point = StringUtils.split(item,",");
			if(point.length == 2)
			{
				int x = new Integer(point[0]).intValue();
				int y = new Integer(point[1]).intValue();
				points.add(new Point(x,y));
			}
		}
		
		return points.toArray(new Point[points.size()]);
	}
	
	

	@Override
	public void ignorableWhitespace(String arg0) throws SAXException
	{
		// TODO Auto-generated method stub
		
	}


	@Override
	public void processingInstruction(String arg0, String arg1)
			throws SAXException
	{
		// TODO Auto-generated method stub
		
	}


	@Override
	public void setDocumentLocator(XLocator arg0) throws SAXException
	{
		// TODO Auto-generated method stub
		
	}


	@Override
	public void startDocument() throws SAXException
	{
		//System.out.println("start");		
	}


	private ShapeType detectShapeType(String parseElement, XAttributeList attribute)
	{
		if (StringUtils.equals(parseElement, "g"))
		{
			int n = attribute.getLength();
			for (short i = 0; i < n; i++)
			{
				String name = attribute.getNameByIndex(i);
				if (StringUtils.equals("class", name))
				{
					String value = attribute.getValueByIndex(i);
					ShapeType [] types = ShapeType.values();
					for(ShapeType type : types)
					{
						if(StringUtils.equals(type.getType(), value))
							return type;
					}
				}
			}
		}
		
		return null;
	}
	
	
	
}
