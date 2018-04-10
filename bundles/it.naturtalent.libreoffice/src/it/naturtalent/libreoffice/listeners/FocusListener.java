package it.naturtalent.libreoffice.listeners;

import com.sun.star.awt.FocusEvent;
import com.sun.star.awt.XFocusListener;
import com.sun.star.awt.XWindow;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;

public class FocusListener implements XFocusListener
{

	private XWindow xWindow;
	
	@Override
	public void disposing(EventObject arg0)
	{
		System.out.println("FocusListener: disposing");

	}

	@Override
	public void focusGained(FocusEvent arg0)
	{		
		//xWindow = UnoRuntime.queryInterface(XWindow.class, arg0.Source);
		System.out.println("FocusListener: gained "+arg0.Source+" Window: "+xWindow);
		//xWindodw.dispose();
		//xWindodw.addMouseListener(new MouseListener());

	}

	@Override
	public void focusLost(FocusEvent arg0)
	{		
		System.out.println("FocusListener: lost "+arg0.FocusFlags);		
		//xWindow.setFocus();


	}

}
