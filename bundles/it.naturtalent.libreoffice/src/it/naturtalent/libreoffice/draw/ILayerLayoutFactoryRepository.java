package it.naturtalent.libreoffice.draw;

import java.util.List;

public interface ILayerLayoutFactoryRepository
{
	// Liest alle Factories aus dem Repository und gibt sie in einer Liste zurueck
	public List<ILayerLayoutFactory>getLayerLayoutFactories();
}
