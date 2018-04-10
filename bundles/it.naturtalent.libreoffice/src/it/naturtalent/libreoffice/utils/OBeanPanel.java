package it.naturtalent.libreoffice.utils;


// OBeanPanel.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, August 2016, 

/* A JPanel envelope for an OOoBean with a loaded read-only
   document.

   Includes mouse and keyboard listeners.


   OOoBean info:
     * source code:   
               https://github.com/LibreOffice/core/blob/
               master/bean/com/sun/star/comp/beans/OOoBean.java 
    
     * JAR location: <OFFICE>\program\classes\officebean.jar

     * Online dev guide:      
               https://wiki.openoffice.org/wiki/Documentation/DevGuide/
               JavaBean/JavaBean_for_Office_Components
    
     * Chapter 16 in Dev Guide, "JavaBean for Office Components"


   Problems with OOoBean:
     https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=63008
           - OOoBean won`t lose its focus
           - connection to OpenOffice is sometimes lost
*/


import javax.swing.*;
import java.awt.event.*;
import java.awt.*;

import com.sun.star.beans.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.frame.*;
import com.sun.star.awt.*;

import com.sun.star.comp.beans.*;

// import com.sun.star.uno.Exception;


public class OBeanPanel extends JPanel
                  implements XMouseClickHandler, XKeyHandler
{
  private int pWidth, pHeight;    // of panel

  private OOoBean oob = null;
  private XComponent doc = null;

  private Font msgFont;
  private FontMetrics fontMetric;
  private String message;


  

  public OBeanPanel(int w, int h)
  {
    pWidth = w;
    pHeight = h;

    setBackground(Color.WHITE);
    setPreferredSize(new Dimension(w, h));
    setLayout(new BorderLayout());

    // start-up and finishing message used by panel
    msgFont = new Font("SansSerif", Font.BOLD, 36);
    fontMetric = getFontMetrics(msgFont);
    message = "Waiting for Office...";

    oob = new OOoBean();   // doesn't connect to Office;
                           // actually an empty method
    add(oob, BorderLayout.CENTER);
  } // end of OBeanPanel()




  public XComponent loadDoc(String fnm)
  // load the document read-only
  {
    PropertyValue[] props = Props.makeProps("ReadOnly", true);
    try  {
      String fileURL = FileIO.fnmToURL(fnm);
      oob.loadFromURL(fileURL, props);  // real work done here!
                // connect to Office, and then load doc
      Lo.setOOoBean(oob);   // initialize my Lo class so I can use
                            // my support classes
      doc = getDoc();
      int docType = Info.reportDocType(doc);

      if (docType == Lo.BASE)
        Lo.dispatchCmd(getFrame(), "DBViewTables", null);    
                                // switch to tables view in Base docs

      XLayoutManager lm = GUI.getLayoutManager(doc);
      lm.setVisible(false);

      // add mouse click & key handlers to the doc
      XUserInputInterception uii = GUI.getUII(doc);
      uii.addMouseClickHandler(this);
      uii.addKeyHandler(this);

      revalidate();   // needs a refresh to appear
      Lo.delay(1000);
    }
    catch(java.lang.Exception e)
    {  System.out.println(e);  }

    return doc;
  }   // end of loadDoc()





  public void paintComponent(Graphics g)
  // display a message while OOoBean is loaded/unloaded
  {
    super.paintComponent(g);
    Graphics2D g2 = (Graphics2D) g;
    g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                        RenderingHints.VALUE_ANTIALIAS_ON);

    int x = (pWidth - fontMetric.stringWidth(message)) / 2;
    int y = (pHeight - fontMetric.getHeight()) / 2;

    g2.setColor(Color.BLUE);
    g2.setFont(msgFont);
    g2.drawString(message, x, y);
  }  // end of paintComponent()




  public void closeDown()
  {
    if (oob != null) {
      System.out.println("Closing connection to office");
      oob.stopOOoConnection();
    }
    Lo.delay(1000);   // wait for close down to finish
    Lo.killOffice();  // make sure office processes are gone
  }  // end of closeDown()



  // --------------- get methods ----------------


  public XComponent getDoc()
  {
    XComponent doc = null;
    try {
      XModel xModel = (XModel)oob.getDocument();
      doc = Lo.qi( XComponent.class, xModel);
    }
    catch(java.lang.Exception e) 
    { System.out.println("OOBean document could not be accessed");  }
    return doc;
  }  // end of getDoc();




  public OOoBean getBean() 
  {  return oob;  }


  public XController getController()
  {
    XController ctrl = null;
    try {
      XModel xModel = (XModel)oob.getDocument();
      ctrl = xModel.getCurrentController();
    }
    catch(java.lang.Exception e) {
      System.out.println("OOBean controller could not be accessed");
    }  
    return ctrl;
  }  // end of getController()




  public XFrame getFrame()
  {
    try {
      return oob.getFrame();
    }
    catch(java.lang.Exception e) {
      System.out.println("OOBean frame could not be accessed");
      return null;
    }  
  }  // end of getFrame()





  // ------------------ XMouseClickHandler methods -------------------

  public boolean mousePressed(com.sun.star.awt.MouseEvent e)
  { System.out.println("Mouse pressed (" + e.X + ", " + e.Y + ")");
    return false;    // send event on, or use true not to send
  }
  
  public boolean mouseReleased(com.sun.star.awt.MouseEvent e)
  { System.out.println("Mouse released (" + e.X + ", " + e.Y + ")");
    return false;
  }
  

  // ---------------------- XKeyHandler methods -------------------------


  @SuppressWarnings("deprecation")
  public boolean keyPressed(com.sun.star.awt.KeyEvent e)
  { 
    // System.out.println("Key pressed: " + e.KeyCode + "/" + e.KeyChar);
    if (oob == null)
      return false;

    try {
      oob.releaseSystemWindow(); 
             // needed to force focus away from bean;
             // must suppress deprecation warning
      oob.aquireSystemWindow();

         // Impress redisplays toolbars; this tries to remove them again
      XLayoutManager lm = GUI.getLayoutManager(doc);
      lm.setVisible(false);

      Lo.dispatchCmd("LeftPaneImpress", 
                  Props.makeProps("LeftPaneImpress", false) );    
         // hide slides pane
      lm.hideElement(GUI.MENU_BAR);
    }
    catch(java.lang.Exception ex) {}
    return false;
  }  // end of keyPressed()


  public boolean keyReleased(com.sun.star.awt.KeyEvent e)
  { return false; }


  // ---------------------- XEventListener methods -------------------------

  public void disposing(com.sun.star.lang.EventObject e){}


} // end of OBeanPanel class
