package it.naturtalent.libreoffice;

import java.util.Map;

import com.sun.star.awt.XWindow;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.UnoRuntime;


import it.naturtalent.libreoffice.draw.DrawDocument;

/**
 * Hilfsfunktionen in Verbindung mit der DrawDesignDatei.
 * 
 * Momentan wird die LibreOffice Draw-Datei fuer die Design-Funktionalitaet benutzt.
 * 
 * 
 * @author dieter
 *
 */
public class DesignHelper
{

	/**
	 * Die von xComponent repraesentierte Drawdatei wird bearbeitbar und im Desktop sichtbar (an oberster Stelle) angezeigt
	 * @param xComponent
	 */
	public static void setFocus(XComponent xComponent)
	{
		XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
		XController xController = xModel.getCurrentController();
		XFrame xframe = xController.getFrame();
		
		XWindow xWindow = xframe.getContainerWindow();
		xWindow.setVisible(true);
		xWindow.setFocus();		
	}
	
}
