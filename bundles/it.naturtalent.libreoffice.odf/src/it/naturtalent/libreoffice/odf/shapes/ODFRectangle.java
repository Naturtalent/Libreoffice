package it.naturtalent.libreoffice.odf.shapes;

public class ODFRectangle
{
	ODFMeasurement x;
	ODFMeasurement y;
	ODFMeasurement width;
	ODFMeasurement height;

	public ODFRectangle(String xmlX, String xmlY, String xmlWidht, String xmlHeight)
	{
		x = new ODFMeasurement();
		x.parseInput(xmlX);
		
		y = new ODFMeasurement();
		y.parseInput(xmlY);

		width = new ODFMeasurement();
		width.parseInput(xmlWidht);
		
		height = new ODFMeasurement();
		height.parseInput(xmlHeight);
	}
	
	@Override
	public String toString()
	{		
		return "X:"+x.toString()+",Y:"+y.toString()+",Width:"+width.toString()+",Height:"+height.toString();
	}

}
