package it.naturtalent.libreoffice.listeners;

import com.sun.star.awt.MouseEvent;
import com.sun.star.awt.Point;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.XGraphics;
import com.sun.star.frame.FeatureStateEvent;
import com.sun.star.frame.XStatusbarController;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XEventListener;
import com.sun.star.uno.Exception;

public class StatusBarListener implements XStatusbarController
{

	@Override
	public void addEventListener(XEventListener arg0)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void dispose()
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void removeEventListener(XEventListener arg0)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void initialize(Object[] arg0) throws Exception
	{
		System.out.println("StatusBarListener initialise");

	}

	@Override
	public void statusChanged(FeatureStateEvent arg0)
	{
		System.out.println("StatusBarListener status changed");

	}

	@Override
	public void disposing(EventObject arg0)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void update()
	{
		System.out.println("StatusBarListener  update");

	}

	@Override
	public void click(Point arg0)
	{
		System.out.println("StatusBarListener  click");

	}

	@Override
	public void command(Point arg0, int arg1, boolean arg2, Object arg3)
	{
		System.out.println("StatusBarListener  command");

	}

	@Override
	public void doubleClick(Point arg0)
	{
		System.out.println("StatusBarListener  doubleclick");

	}

	@Override
	public boolean mouseButtonDown(MouseEvent arg0)
	{
		System.out.println("StatusBarListener  button down");
		return false;
	}

	@Override
	public boolean mouseButtonUp(MouseEvent arg0)
	{
		System.out.println("StatusBarListener  buttonup");
		return false;
	}

	@Override
	public boolean mouseMove(MouseEvent arg0)
	{
		System.out.println("StatusBarListener mousemove");
		return false;
	}

	@Override
	public void paint(XGraphics arg0, Rectangle arg1, int arg2)
	{
		// TODO Auto-generated method stub

	}

}
