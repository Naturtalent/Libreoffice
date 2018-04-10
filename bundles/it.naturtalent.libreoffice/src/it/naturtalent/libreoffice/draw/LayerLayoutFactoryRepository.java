package it.naturtalent.libreoffice.draw;

import java.util.ArrayList;
import java.util.List;

/**
 * Diese Klasse realisiert das Repository indem alle Factories gespeichet werden.
 * 
 * @author dieter
 *
 */
public class LayerLayoutFactoryRepository implements ILayerLayoutFactoryRepository
{
	private static List<ILayerLayoutFactory> layerLayoutFactory = new ArrayList<ILayerLayoutFactory>();

	@Override
	public List<ILayerLayoutFactory> getLayerLayoutFactories()
	{		
		return layerLayoutFactory;
	}
}
