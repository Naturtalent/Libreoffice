package it.naturtalent.libreoffice.listeners;

import org.eclipse.e4.core.services.events.IEventBroker;
import org.jnativehook.mouse.NativeMouseEvent;
import org.jnativehook.mouse.NativeMouseListener;

import it.naturtalent.libreoffice.DrawDocumentEvent;

public class GlobalMouseListener implements NativeMouseListener
{

	private IEventBroker eventBroker;
	
	/**
	 * Die Implementierung des globalen MouseListeners.
	 * Das empfangene Event wird mit dem EventBroker weitergegeben
	 * @param eventBroker
	 */
	public GlobalMouseListener(IEventBroker eventBroker)
	{
		super();
		this.eventBroker = eventBroker;
	}

	@Override
	public void nativeMouseClicked(NativeMouseEvent arg0)
	{
		//System.out.println("GlobalMouse - clicked");
	}

	@Override
	public void nativeMousePressed(NativeMouseEvent arg0)
	{
		//System.out.println("GlobalMouse - pressed");
		eventBroker.post(DrawDocumentEvent.DRAWDOCUMENT_EVENT_GLOBALMOUSEPRESSED, arg0);
	}

	@Override
	public void nativeMouseReleased(NativeMouseEvent arg0)
	{
		//System.out.println("GlobalMouse - release");

	}

}
