package it.naturtalent.libreoffice.listeners;

import com.sun.star.awt.MouseEvent;
import com.sun.star.awt.XMouseListener;
import com.sun.star.lang.EventObject;

public class MouseListener implements XMouseListener
{

	@Override
	public void disposing(EventObject arg0)
	{
		System.out.println("XMouseListener: disposing");

	}

	@Override
	public void mouseEntered(MouseEvent arg0)
	{
		System.out.println("XMouseListener: enter");

	}

	@Override
	public void mouseExited(MouseEvent arg0)
	{
		System.out.println("XMouseListener: exit");

	}

	@Override
	public void mousePressed(MouseEvent arg0)
	{
		System.out.println("XMouseListener: pressed");
	}

	@Override
	public void mouseReleased(MouseEvent arg0)
	{
		System.out.println("XMouseListener: release");

	}

}
