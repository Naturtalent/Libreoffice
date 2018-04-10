package it.naturtalent.libreoffice.draw;

import it.naturtalent.libreoffice.math.BigDecimalMath;

import java.math.BigDecimal;

import org.eclipse.swt.graphics.Rectangle;

import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.HomogenMatrix3;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;

public class LineShape extends Shape
{
	public LineShape(XShape xShape)
	{
		super(xShape);		
	}

	public LineShape(Layer layer, Rectangle bound) throws Exception
	{
		super(layer, bound, ShapeType.LineShape);		
	}

	/*
	@Override
	public String getShapeLabel()
	{
		return ShapeType.LineShape.name();
	}
	*/
	
	public BigDecimal getLength()
	{
		XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(
				XPropertySet.class, xShape);
		try
		{
			HomogenMatrix3 aHomogenMatrix3 = (HomogenMatrix3) xPropSet
					.getPropertyValue("Transformation");

			BigDecimal breit = new BigDecimal(aHomogenMatrix3.Line1.Column1);
			BigDecimal  hoch = new BigDecimal(aHomogenMatrix3.Line2.Column2);	
			
			breit = BigDecimalMath.powRound(breit,2);
			hoch = BigDecimalMath.powRound(hoch,2);
			return BigDecimalMath.sqrt(breit.add(hoch ));
			
		} catch (UnknownPropertyException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WrappedTargetException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

}
