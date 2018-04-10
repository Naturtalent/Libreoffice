package it.naturtalent.libreoffice.odf.shapes;

public class ODFPoint
{
	ODFMeasurement x;
	ODFMeasurement y;
	
	public ODFPoint(String xmlX, String xmlY)
	{
		x = new ODFMeasurement();
		x.parseInput(xmlX);
		
		y = new ODFMeasurement();
		y.parseInput(xmlY);
	}
	
	@Override
	public String toString()
	{		
		return "X:"+x.toString()+",y:"+y.toString();
	}
}
