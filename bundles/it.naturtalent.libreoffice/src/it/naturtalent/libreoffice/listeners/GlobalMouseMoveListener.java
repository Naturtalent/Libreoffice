package it.naturtalent.libreoffice.listeners;

import org.eclipse.e4.core.services.events.IEventBroker;
import org.jnativehook.mouse.NativeMouseEvent;
import org.jnativehook.mouse.NativeMouseInputListener;

import it.naturtalent.libreoffice.DrawDocumentEvent;
import it.naturtalent.libreoffice.draw.DrawDocument;

public class GlobalMouseMoveListener implements NativeMouseInputListener
{
	private IEventBroker eventBroker;
	
	
	/**
	 * Konstruktion
	 * 
	 * @param eventBroker
	 */
	public GlobalMouseMoveListener(IEventBroker eventBroker)
	{
		super();
		this.eventBroker = eventBroker;
	}

	@Override
	public void nativeMouseClicked(NativeMouseEvent arg0)
	{
		System.out.println("GlobalMouse - clicked");
		//eventBroker.post(DrawDocumentEvent.DRAWDOCUMENT_EVENT_GLOBALMOUSECLICK, arg0);

	}

	@Override
	public void nativeMousePressed(NativeMouseEvent arg0)
	{
		System.out.println("GlobalMouse - pressed");
		eventBroker.post(DrawDocumentEvent.DRAWDOCUMENT_EVENT_GLOBALMOUSEPRESSED, arg0);
	}

	@Override
	public void nativeMouseReleased(NativeMouseEvent arg0)
	{
		//System.out.println("GlobalMouse - release");

	}

	@Override
	public void nativeMouseDragged(NativeMouseEvent arg0)
	{
		//System.out.println("GlobalMouse - dragged");
		
	}

	@Override
	public void nativeMouseMoved(NativeMouseEvent arg0)
	{
		//System.out.println("GlobalMouse - moved");
		
	}

}
