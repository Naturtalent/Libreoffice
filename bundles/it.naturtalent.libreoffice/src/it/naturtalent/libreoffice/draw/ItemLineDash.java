package it.naturtalent.libreoffice.draw;

import java.beans.PropertyChangeEvent;



/**
 * Datenstruktur Linienstyle (gestrichelte, gepunktete Linien)
 * 
 * @author dieter
 *
 */
public class ItemLineDash extends BaseBean
{
	public static final String PROP_LINEDASHDOTS = "linedashdots"; //$NON-NLS-N$
	public static final String PROP_LINEDASHDOTLEN = "linedashdotlen"; //$NON-NLS-N$
	public static final String PROP_LINEDASHES = "linedashes"; //$NON-NLS-N$
	public static final String PROP_LINEDASHLEN = "linedashlen"; //$NON-NLS-N$
	public static final String PROP_LINEDASHDISTANCE = "linedashdistance"; //$NON-NLS-N$
	
	private Short dots;			// Anzahl der Punkte
	private Integer dotlen;		// Laenge eines Punktes
	private Short dashes;		// Anzahl der Striche
	private Short dashlen;		// Laenge eines einzelnen Striches
	private Integer distance;		// Distance zwischen den Punkten

	public Short getDots()
	{
		return dots;
	}
	public void setDots(Short dots)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINEDASHDOTS, this.dots, this.dots = dots));		
	}
	
	public Integer getDotlen()
	{
		return dotlen;
	}
	public void setDotlen(Integer dotlen)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINEDASHDOTLEN, this.dotlen, this.dotlen = dotlen));
	}
	
	public Short getDashes()
	{
		return dashes;
	}
	public void setDashes(Short dashes)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINEDASHES, this.dashes, this.dashes = dashes));		
	}
	
	public Short getDashlen()
	{
		return dashlen;
	}
	public void setDashlen(Short dashlen)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINEDASHLEN, this.dashlen, this.dashlen = dashlen));
	}
	
	public Integer getDistance()
	{
		return distance;
	}
	public void setDistance(Integer distance)
	{
		firePropertyChange(new PropertyChangeEvent(this, PROP_LINEDASHDISTANCE, this.distance, this.distance = distance));		
	}
	
	/*
	public LineDash getLineDash()
	{
		LineDash aLineDash = new LineDash();
		aLineDash.Dashes = dashes;
		aLineDash.DashLen = dashlen;
		aLineDash.Distance = distance;
		aLineDash.Dots = dots;
		aLineDash.DotLen = dotlen;
		return aLineDash;
	}
	*/
	
	@Override
	public int hashCode()
	{
		final int prime = 31;
		int result = 1;
		result = prime * result + ((dashes == null) ? 0 : dashes.hashCode());
		result = prime * result + ((dashlen == null) ? 0 : dashlen.hashCode());
		result = prime * result
				+ ((distance == null) ? 0 : distance.hashCode());
		result = prime * result + ((dotlen == null) ? 0 : dotlen.hashCode());
		result = prime * result + ((dots == null) ? 0 : dots.hashCode());
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
		ItemLineDash other = (ItemLineDash) obj;
		if (dashes == null)
		{
			if (other.dashes != null)
				return false;
		}
		else if (!dashes.equals(other.dashes))
			return false;
		if (dashlen == null)
		{
			if (other.dashlen != null)
				return false;
		}
		else if (!dashlen.equals(other.dashlen))
			return false;
		if (distance == null)
		{
			if (other.distance != null)
				return false;
		}
		else if (!distance.equals(other.distance))
			return false;
		if (dotlen == null)
		{
			if (other.dotlen != null)
				return false;
		}
		else if (!dotlen.equals(other.dotlen))
			return false;
		if (dots == null)
		{
			if (other.dots != null)
				return false;
		}
		else if (!dots.equals(other.dots))
			return false;
		return true;
	}


	
}
