package it.naturtalent.libreoffice.draw;

import java.beans.PropertyChangeEvent;





/**
 * Datenstruktur der definierten Styleeigenschaften
 * 
 * @author dieter
 *
 */
public class ItemStyle extends BaseBean implements IItemStyle, Cloneable
{

	// Properties
	public static final String PROP_LINESTYLENAME = "linestylename"; //$NON-NLS-N$	
	public static final String PROP_LINECOLOR = "linecolor"; //$NON-NLS-N$
	public static final String PROP_LINEWIDTH = "linewidth"; //$NON-NLS-N$
	public static final String PROP_LINEDASH = "linedash"; //$NON-NLS-N$
	
	
	// Name
	private String name;			// Stylename
	private Integer lineColor;		// Farbe der Linie
	private Integer lineWidth;		// Breite der Linie
	private ItemLineDash linedash;
	
	// Linie
	
	public String getName()
	{
		return name;
	}
	public void setName(String name)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINESTYLENAME, this.name, this.name = name));		
	}
	
	public Integer getLineColor()
	{
		return lineColor;
	}
	public void setLineColor(Integer lineColor)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINECOLOR, this.lineColor, this.lineColor = lineColor));		
	}
	
	public Integer getLineWidth()
	{
		return lineWidth;
	}
	public void setLineWidth(Integer lineWidth)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINEWIDTH, this.lineWidth, this.lineWidth = lineWidth));
	}
	
	public ItemLineDash getLinedash()
	{
		return linedash;
	}
	public void setLinedash(ItemLineDash linedash)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINEDASH, this.linedash, this.linedash = linedash));		
	}
	
	@Override
	public int hashCode()
	{
		final int prime = 31;
		int result = 1;
		result = prime * result
				+ ((lineColor == null) ? 0 : lineColor.hashCode());
		result = prime * result
				+ ((lineWidth == null) ? 0 : lineWidth.hashCode());
		result = prime * result
				+ ((linedash == null) ? 0 : linedash.hashCode());
		result = prime * result + ((name == null) ? 0 : name.hashCode());
		return result;
	}
	@Override
	public boolean equals(Object obj)
	{
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		ItemStyle other = (ItemStyle) obj;
		if (lineColor == null)
		{
			if (other.lineColor != null)
				return false;
		}
		else if (!lineColor.equals(other.lineColor))
			return false;
		if (lineWidth == null)
		{
			if (other.lineWidth != null)
				return false;
		}
		else if (!lineWidth.equals(other.lineWidth))
			return false;
		if (linedash == null)
		{
			if (other.linedash != null)
				return false;
		}
		else if (!linedash.equals(other.linedash))
			return false;
		if (name == null)
		{
			if (other.name != null)
				return false;
		}
		else if (!name.equals(other.name))
			return false;
		return true;
	}
	
	
}
