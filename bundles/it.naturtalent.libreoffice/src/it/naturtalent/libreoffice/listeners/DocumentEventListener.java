package it.naturtalent.libreoffice.listeners;

import com.sun.star.document.DocumentEvent;
import com.sun.star.document.XDocumentEventBroadcaster;
import com.sun.star.document.XDocumentEventListener;
import com.sun.star.lang.EventObject;
import com.sun.star.uno.UnoRuntime;

/**
 * @author dieter
 *
 *
 *	XDocumentEventBroadcaster docEventBroadcaster = UnoRuntime.queryInterface(
 *	XDocumentEventBroadcaster.class, xComponent );			
 *	XocEventBroadcaster.addDocumentEventListener( new XDocumentEventListener()
 *
 *
 */
public class DocumentEventListener implements XDocumentEventListener
{

	@Override
	public void disposing(EventObject arg0)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void documentEventOccured(DocumentEvent arg0)
	{
		System.out.println("DOCUMENT EVENT LISTENER");
	}

}
