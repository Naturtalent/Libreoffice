package it.naturtalent.libreoffice;

import java.io.File;

import org.eclipse.swt.graphics.Rectangle;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameContainer;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;

import it.naturtalent.libreoffice.draw.DrawDocument;
import it.naturtalent.libreoffice.draw.Shape;

public class ImportGraphicShape
{

	private DrawDocument drawDocument;

	public ImportGraphicShape(DrawDocument drawDocument)
	{
		super();
		this.drawDocument = drawDocument;
	}
	
	public Shape readShape(File sourceFile, Rectangle bound)
	{
		Shape shape = null;
		
		try
		{
			XMultiServiceFactory xFactory = UnoRuntime.queryInterface(
					XMultiServiceFactory.class, drawDocument.getxComponent());
			
			Object graphicShape = xFactory.createInstance( "com.sun.star.drawing.GraphicObjectShape");			
			XShape xShape = (XShape)UnoRuntime.queryInterface(XShape.class, graphicShape);
			
			shape = new Shape(xShape);			
			xShape.setSize(new Size(bound.width, bound.height));
			xShape.setPosition(new Point(bound.x, bound.y));
			
			  // Creating bitmap container service
            XNameContainer bitmapContainer = UnoRuntime.queryInterface(
                              XNameContainer.class,
                              xFactory.createInstance(
                                   "com.sun.star.drawing.BitmapTable"));
            
            // Inserting test image to the container
            StringBuffer sourceURL = new StringBuffer("file:///");
            sourceURL.append(sourceFile.getCanonicalPath().replace('\\', '/')); 
            bitmapContainer.insertByName("testimg",sourceURL.toString());
           
            
            // Querying property interface for the graphic shape service
            XPropertySet xPropSet = (XPropertySet)UnoRuntime.queryInterface(
                                  XPropertySet.class, graphicShape);
            
            // Assign test image internal URL to the graphic shape property
            xPropSet.setPropertyValue("GraphicURL",bitmapContainer.getByName("testimg"));            
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		 return shape;
	}
	
}
