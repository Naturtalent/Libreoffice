package it.naturtalent.libreoffice.draw;

/**
 * Interface beschreibt das erweiterte Verhalten eines Layers.
 * 
 * Hierzu gehoeren;
 * Ausgabe des Shapes in die Ebene des DrawDocuments
 * Auswertung aller Shapes und Anzeige in der DetailsView
 * 
 * @author dieter
 *
 */
public interface ILayerLayout
{

	// Name des Layouts (identisch mit dem Model Layer und der Ebene im DrawDocument)
	public String getName ();
	
	public void setName(String name);
	
	// fuegt ein Shape (ggf. auswaehlbar) in dem DrawDocument Layer an der, aus der Statusbar ausgelesenen Mouseposition,  ein
	public void addShape();
	
	// die durch die Shapes reprasentierten Daten in einem Detailfenster anzeigen
	public void showDetails();
}
