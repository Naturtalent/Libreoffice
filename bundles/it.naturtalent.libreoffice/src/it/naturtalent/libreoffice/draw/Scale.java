package it.naturtalent.libreoffice.draw;

import java.math.BigDecimal;

import it.naturtalent.libreoffice.ShapeHelper;
import it.naturtalent.libreoffice.Utils;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XInterface;

public class Scale
{
	
	// MeasurementUnits (Auswahl http://api.libreoffice.org/docs/idl/ref/MeasureUnit_8idl.html)
	public final static Short MEASUREUNIT_MM_100TH = 0; 	// hundertstel mm
	public final static Short MEASUREUNIT_MM_10TH = 1;	 	// zehntel mm
	public final static Short MEASUREUNIT_MM = 2;		 	// mm
	public final static Short MEASUREUNIT_CM = 3;	 		// cm
	public final static Short MEASUREUNIT_M = 10;	 		// m
	
	// Massstab Zaehler und Nenner
	private Integer scaleNumerator = 1;
	private Integer scaleDenominator = 50;
	
	// Die Masseinheit
	private Short measureUnit = MEASUREUNIT_M;
	
	// Faktor (m) (Shapes werden original in 1/100 mm gespeichert)
	public static BigDecimal measureFactor = new BigDecimal(100000.0);
	
	
	private XComponent xComponent;
	
	/**
	 * Konstruktion
	 * 
	 * @param xComponent
	 */
	public Scale(XComponent xComponent)
	{
		super();
		this.xComponent = xComponent;
	}

	/**
	 * Die Scaledaten an die Komponente uebertragen.
	 * 
	 */
	public void pushScaleProperties()
	{
		if (xComponent != null)
		{
			try
			{
				XMultiServiceFactory xFactory = UnoRuntime.queryInterface(
						XMultiServiceFactory.class, xComponent);
				XInterface settings = (XInterface) xFactory
						.createInstance("com.sun.star.drawing.DocumentSettings");
				XPropertySet xPageProperties = UnoRuntime.queryInterface(
						XPropertySet.class, settings);
				
				xPageProperties.setPropertyValue("ScaleNumerator",scaleNumerator);
				xPageProperties.setPropertyValue("ScaleDenominator",scaleDenominator);
				xPageProperties.setPropertyValue("MeasureUnit", measureUnit);
				
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	/**
	 * Die Scaledaten von der Komponente lesen.
	 * 
	 */
	public void pullScaleProperties()
	{
		if (xComponent != null)
		{
			try
			{
				XMultiServiceFactory xFactory = UnoRuntime.queryInterface(
						XMultiServiceFactory.class, xComponent);
				XInterface settings = (XInterface) xFactory
						.createInstance("com.sun.star.drawing.DocumentSettings");
				XPropertySet xPageProperties = UnoRuntime.queryInterface(
						XPropertySet.class, settings);
				scaleNumerator = (Integer) xPageProperties
						.getPropertyValue("ScaleNumerator");
				scaleDenominator = (Integer) xPageProperties
						.getPropertyValue("ScaleDenominator");

				measureUnit = (Short) xPageProperties
						.getPropertyValue("MeasureUnit");
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	public void setScaleDenominator(Integer scaleDenominator)
	{
		this.scaleDenominator = scaleDenominator;		
	}
	
	public Integer getScaleDenominator()
	{
		return scaleDenominator;
	}
	
	public Short getMeasureUnit()
	{
		return measureUnit;
	}

	public void setMeasureUnit(Short measureUnit)
	{
		this.measureUnit = measureUnit;
	}
	
	public Point scalePoint(Point pt)
	{
		double scaleFactor = ((double)scaleDenominator/(double)scaleNumerator);
		
		pt.X = (int) (pt.X * scaleFactor);
		pt.Y = (int) (pt.Y * scaleFactor);
		
		return pt;
	}

	/**
	 * Der Devisor wird bestimmt indem eine Referenzlaenge (z.B. eine gezeichnete Linie) durch einen 
	 * Zielwert geteilt wird. Man definiert: die gezeichnet Linie hat eine Laenge von x-Meter und erhaelt damit den
	 * Masstab.
	 * 
	 * @param lineLength
	 * @param dialogLength
	 */
	public void calculateScaleByLineLength(BigDecimal lineLength, BigDecimal dialogLength)
	{
		BigDecimal doubleMass = dialogLength.multiply(measureFactor);		
		doubleMass = doubleMass.divide(lineLength, BigDecimal.ROUND_UP);
		scaleDenominator = doubleMass.intValue();
		pushScaleProperties();
	}
	
	public void createRefenceShape()
	{
        try
		{
			XShape referenceLine = ShapeHelper.createShape( xComponent,
			        new Point( 1000, 1000 ), new Size( 5000, 5000 ),
			            "com.sun.star.drawing.RectangleShape" );
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
}
