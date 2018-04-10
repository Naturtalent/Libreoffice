package it.naturtalent.libreoffice.draw;

public interface ILayerLayoutFactory
{
	// Name der LayerLayouts, die mit dieser Factory erzeugt werden.
	public String getName();
	
	public ILayerLayout createLayout();
}
