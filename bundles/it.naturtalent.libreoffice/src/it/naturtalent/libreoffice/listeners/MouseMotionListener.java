package it.naturtalent.libreoffice.listeners;

import com.sun.star.awt.MouseEvent;
import com.sun.star.awt.XMouseMotionListener;
import com.sun.star.lang.EventObject;

public class MouseMotionListener implements XMouseMotionListener
{

	@Override
	public void disposing(EventObject arg0)
	{
		System.out.println("XMouseMotionListener: disposing");
	}

	@Override
	public void mouseDragged(MouseEvent arg0)
	{
		System.out.println("XMouseMotionListener: dragged");

	}

	@Override
	public void mouseMoved(MouseEvent arg0)
	{
		System.out.println("XMouseMotionListener: moved");

	}

}
