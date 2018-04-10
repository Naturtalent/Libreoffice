package it.naturtalent.libreoffice.draw;


public interface IItemStyle
{
	
	/**
	 * Liefert den Namen dieser Datenstruktur:
	 * @return
	 */
	public String getName();

	/**
	 * Definiert den Namen dieser Datenstuktur.
	 * 
	 * @param name
	 */
	public void setName(String name);
	
	
	/**
	 * Liefert die Linienstaerke.
	 * 
	 * @return
	 */
	public Integer getLineWidth();
		
	/**
	 * Definiert die Linientaerke.
	 * 
	 * @param lineWidth
	 */
	public void setLineWidth(Integer lineWidth);
	
	/**
	 * Liefert die Linienfarbe.
	 * 
	 * @return
	 */
	public Integer getLineColor();
		
	/**
	 * Definiert die Linienfarbe.
	 * 
	 * @param lineColor
	 */
	public void setLineColor(Integer lineColor);
	
	
	/**
	 * Rueckgabe des Linienstyles.
	 * 
	 * @return
	 */
	public ItemLineDash getLinedash();
		
	/**
	 * Definiert den Linienstyle-
	 * 
	 * @param linedash
	 */
	public void setLinedash(ItemLineDash linedash);
	
}
