package it.naturtalent.libreoffice.odf.shapes;

import java.util.LinkedList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;

public class ODFMeasurement
{

	Double value;
	String measureUnit;
	private int lastIdx;
	
	/**
	 * Eingabestring in der Form '1.23cm'
	 * @param xmlMeasurement
	 */
	public void parseInput(String xmlMeasurement)
	{
		measureUnit = null;		
		if (StringUtils.isNotEmpty(xmlMeasurement))
		{
			LinkedList<String> numbers = new LinkedList<String>();
			Pattern p = Pattern.compile("\\d+");
			Matcher m = p.matcher(xmlMeasurement);
			while (m.find())
			{
				numbers.add(m.group());
				lastIdx = m.end();
			}
			String number = numbers.get(0);

			if (numbers.size() > 1)
				number = number + "." + numbers.get(1);
			value = new Double(number);

			measureUnit = StringUtils.substring(xmlMeasurement, lastIdx);
		}
	};
	
	@Override
	public String toString()
	{		
		return value.toString()+measureUnit;
	}
}
