package it.naturtalent.libreoffice.draw;

import org.apache.commons.lang3.StringUtils;
import org.eclipse.swt.graphics.Rectangle;

import it.naturtalent.libreoffice.ShapeHelper;
import it.naturtalent.libreoffice.Utils;
import it.naturtalent.libreoffice.draw.Shape.ShapeType;

import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;

public class TextShape extends Shape
{

	public TextShape(XShape xShape)
	{
		super(xShape);		
	}
	
	public TextShape(Layer layer, Rectangle bound) throws Exception
	{
		super(layer, bound, ShapeType.TextShape);		
	}
	
	
	public String getText()
	{
		XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(
				XPropertySet.class, xShape);
		
		try
		{
			String text = (String) xPropSet.getPropertyValue("UINameSingular");
			if(StringUtils.isNotEmpty(text))
			{
				String [] split = StringUtils.split(text,'\'');
				return split[1];
			}
			
		} catch (UnknownPropertyException | WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return null;
	}
	
	public void setText(String text)
	{
		ShapeHelper.addPortion(xShape, text, false);
	}


}
