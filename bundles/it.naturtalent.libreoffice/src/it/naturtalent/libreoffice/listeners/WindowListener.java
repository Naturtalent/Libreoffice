package it.naturtalent.libreoffice.listeners;

import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XWindowListener;
import com.sun.star.lang.EventObject;

public class WindowListener implements XWindowListener
{

	@Override
	public void disposing(EventObject arg0)
	{
		System.out.println("WindowListener: disposing");
	}

	@Override
	public void windowHidden(EventObject arg0)
	{
		System.out.println("WindowListener: hidden");

	}

	@Override
	public void windowMoved(WindowEvent arg0)
	{
		System.out.println("WindowListener: moved");
	}

	@Override
	public void windowResized(WindowEvent arg0)
	{
		System.out.println("WindowListener: resize");

	}

	@Override
	public void windowShown(EventObject arg0)
	{
		System.out.println("WindowListener: shown");

	}

}
