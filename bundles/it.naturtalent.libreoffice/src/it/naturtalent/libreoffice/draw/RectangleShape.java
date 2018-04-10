package it.naturtalent.libreoffice.draw;

import org.eclipse.swt.graphics.Rectangle;

public class RectangleShape extends Shape
{
	public RectangleShape(Layer layer, Rectangle bound) throws Exception
	{
		super(layer, bound, ShapeType.RectangleShape);		
	}

}
