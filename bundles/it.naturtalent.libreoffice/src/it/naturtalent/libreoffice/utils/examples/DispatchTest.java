package it.naturtalent.libreoffice.utils.examples;


// DispatchTest.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, March 2015

/* Shows the use of Robot and the dispatcher

   A Help window is opened (in a browser), and a slide presentation
   is started. You must close the presentation and Office at the end.

   Usage: 
       > compile DispatchText.java
       > run DispatchTest <fnm>
    e.g.
       > run DispatchTest algs.odp

*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.view.*;

import it.naturtalent.libreoffice.utils.GUI;
import it.naturtalent.libreoffice.utils.Lo;

import com.sun.star.beans.*;

import java.awt.Robot;
import java.awt.AWTException;
import java.awt.event.KeyEvent;


public class DispatchTest
{



  public DispatchTest(String fnm)
  {
    XComponentLoader loader = Lo.loadOffice();   

    XComponent doc = Lo.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);
    Lo.delay(100);
    // toggleSlidePane();

    Lo.dispatchCmd("HelpIndex");    // .uno:HelpIndex

    Lo.dispatchCmd("Presentation");   // .uno:Presentation

    //Lo.closeDoc(doc);
    //Lo.closeOffice();
  }  // end of DispatchTest()



  private void toggleSlidePane()
  // alt-v and then alt-l to foreground window
  // makes slide pane appear/disappear in Impress
  {
    try {
      Robot robot = new Robot();
      robot.setAutoDelay(250);
      robot.keyPress(KeyEvent.VK_ALT);

      robot.keyPress(KeyEvent.VK_V);
      robot.keyRelease(KeyEvent.VK_V);

      robot.keyPress(KeyEvent.VK_L);
      robot.keyRelease(KeyEvent.VK_L);

      robot.keyRelease(KeyEvent.VK_ALT);
    }
    catch(AWTException e)
    {  System.out.println("sendkeys slidePane exception: " + e); }
  }   // end of toggleSlidePane()




  // -----------------------------------------------


  public static void main(String args[])
  {  
    if (args.length == 1) 
       new DispatchTest(args[0]); 
    else
      System.out.println("Usage: DispatchTest fnm");
  }  // end of main()


}  // end of DispatchTest class
