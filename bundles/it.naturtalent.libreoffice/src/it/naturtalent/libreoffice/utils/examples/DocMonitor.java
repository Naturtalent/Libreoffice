package it.naturtalent.libreoffice.utils.examples;


// DocMonitor.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, March 2015

/* Monitor creation of office bridge, and office closing.
   The remote bridge is only possible if office is created with a socket.

   Usage: 
       > run DocMonitor <fnm>
    e.g.
       > run DocMonitor algs.odp

   Useful tools:
       > lolist
       > lokill
    
       > netstat | grep 8100
            // for when office is listening on socket at port 8100
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.view.*;

import it.naturtalent.libreoffice.utils.GUI;
import it.naturtalent.libreoffice.utils.Lo;

import com.sun.star.beans.*;

import com.sun.star.connection.*;
import com.sun.star.bridge.*;

import com.sun.star.awt.*;


public class DocMonitor
{

  public DocMonitor(String fnm)
  {
    XComponentLoader loader =  Lo.loadOffice();   
                              //Lo.loadSocketOffice();

    XDesktop xDesktop = Lo.getDesktop();
    xDesktop.addTerminateListener( new XTerminateListener()
    {
       public void queryTermination(EventObject e) throws TerminationVetoException
       {  System.out.println("TL: Starting Closing");   } 

       public void notifyTermination(EventObject e)
       {  System.out.println("TL: Finished Closing"); } 
       
       public void disposing(EventObject e) 
       {  System.out.println("TL: Disposing");  }
    });


    XComponent bridgeComp = Lo.getBridge();
    if (bridgeComp != null) {
      System.out.println("Found bridge");
      bridgeComp.addEventListener( new XEventListener()
      {
        public void disposing(EventObject e)
        { /* remote bridge has gone down, because the 
              office crashed or was terminated. */
          System.out.println("Office bridge has gone!!");
          System.exit(1);
        }
      });
    }

    XComponent doc = Lo.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);

    System.out.println("Waiting for 5 secs before closing doc...");
    Lo.delay(5000);
    Lo.closeDoc(doc);

    System.out.println("Waiting for 5 secs before closing Office...");
    Lo.delay(5000);
    Lo.closeOffice();
  }  // end of DocMonitor()


  // -----------------------------------------------


  public static void main(String args[])
  {  
    if (args.length == 1) 
       new DocMonitor(args[0]); 
    else
      System.out.println("Usage: DocMonitor fnm");
  }  // end of main()


}  // end of DocMonitor class
