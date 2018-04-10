package it.naturtalent.libreoffice.utils.examples;


// DocWindow.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, March 2015

/* Monitor the stages of window creation, activation, minimizing, 
   restoring, closing...

   Usage:
     compile DocWindow.java

     > run DocWindow <fnm>
     e.g.
     > run DocWindow algs.odp
*/

import com.sun.star.awt.*;
import com.sun.star.uno.*;

import it.naturtalent.libreoffice.utils.GUI;
import it.naturtalent.libreoffice.utils.JNAUtils;
import it.naturtalent.libreoffice.utils.Lo;

import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.jna.platform.win32.WinDef.HWND;



public class DocWindow implements XTopWindowListener
{

  public DocWindow(String fnm)
  {
    XComponentLoader loader = Lo.loadOffice();

    XExtendedToolkit tk = Lo.createInstanceMCF(XExtendedToolkit.class, 
                                       "com.sun.star.awt.Toolkit");
    if (tk != null)
      tk.addTopWindowListener(this);

    XComponent doc = Lo.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);
    // triggers 2 opened and 2 activated events

    System.out.println("\nWaiting...");
    Lo.delay(3000);

    // System.out.println("\nWindow Handle integer: " + GUI.getWindowHandle(doc));

    HWND handle = JNAUtils.getHandle();
    System.out.println("\nHandle: \"" + JNAUtils.handleString(handle) + "\"");

    System.out.println("\nMinimizing...");
    JNAUtils.winMinimize(handle);  // triggers minimized and de-activated events
    Lo.delay(3000);

    System.out.println("\nRestoring...");
    JNAUtils.winRestore(handle);  // triggers normalized and activated events
    Lo.delay(3000);

    Lo.closeDoc(doc);   // triggers de-activated and 2 closed events (no closing)
    Lo.delay(3000);
    Lo.closeOffice();
  } // end of DocWindow()



  // XTopWindowListener methods

  public void windowOpened(EventObject event) 
  { 
    System.out.println("WL: Opened");  
    XWindow xWin =  UnoRuntime.queryInterface(XWindow.class, event.Source);
    GUI.printRect( xWin.getPosSize());  
  }  // end of windowOpened()


  public void windowActivated(EventObject event)
  { System.out.println("WL: Activated");  
    System.out.println("  Title bar: \"" + GUI.getTitleBar() + "\"");
  }  // end of windowActivated()


  public void windowMinimized(EventObject event)
  { System.out.println("WL: Minimized");  }

  public void windowNormalized(EventObject event)
  { System.out.println("WL: Normalized");  }

  public void windowDeactivated(EventObject event)
  { System.out.println("WL: De-activated");  }

  public void windowClosing(EventObject event)   // never called
  { System.out.println("WL: Closing");  }

  public void windowClosed(EventObject event)
  { System.out.println("WL: Closed");  }

  public void disposing(EventObject event)   // never called
  { System.out.println("WL: Disposing");  }



  // --------------------------------------------------


  public static void main(String args[])
  {  
    if (args.length == 1) 
       new DocWindow(args[0]); 
    else
      System.out.println("Usage: DocWindow fnm");
  }  // end of main()

}  // end of DocWindow class

