package it.naturtalent.libreoffice;

import com.sun.star.awt.XWindow;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.TerminationVetoException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XTerminateListener;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;

import it.naturtalent.libreoffice.utils.GUI;
import it.naturtalent.libreoffice.utils.Lo;

/**
 * 
 * 
 * @author Dieter Apel
 *
 */
public class OpenLoDocument
{

	private static XComponentLoader officeDocumentLoader = null;

	/**
	 * @param documentPath
	 */
	public static void loadLoDocument(final String documentPath)
	{

		if (officeDocumentLoader == null)
			officeDocumentLoader = Lo.getOfficeLoader();

		if (officeDocumentLoader != null)
		{			
			// Listener ueberwacht 'Libreoffice beenden'
			XDesktop xDesktop = Lo.getDesktop();
			xDesktop.addTerminateListener(new XTerminateListener()
			{
				public void queryTermination(EventObject e)
						throws TerminationVetoException
				{
					//System.out.println("TL: Starting Closing");
				}

				public void notifyTermination(EventObject e)
				{
					// Lo beendet - Loader ist nicht mehr gueltig
					officeDocumentLoader = null;
					//System.out.println("XDesktop terminated");
				}

				public void disposing(EventObject e)
				{
					//System.out.println("TL: Disposing");										
				}
			});
			
			// das Dokument wird geladen
			XComponent xDocument = Lo.openDoc(documentPath,officeDocumentLoader);
			if (xDocument != null)
			{
				if (xDocument != null)
				{
					// ContainerWindow (XWindow) wird sichtbar und erhaelt den
					// Focos
					GUI.setVisible(xDocument, true);

					// Listener ueberwacht das Schliessen des Dokuments
					XWindow xWindow = GUI.getFrame(xDocument).getContainerWindow();
					if (xWindow != null)
					{
						xWindow.addEventListener(new XEventListener()
						{
							@Override
							public void disposing(EventObject arg0)
							{
								Object obj = arg0.Source;
								
								XPropertySet props = Lo.qi(XPropertySet.class, obj);
								  
								
								System.out.println("Containerwindow disposed");
							}
						});
					}
				}
			}
		}
	}
	

}
