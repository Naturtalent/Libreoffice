package it.naturtalent.libreoffice.utils.examples;


// DocProps.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Feb. 2017

/*  Print document properties
    (e.g. author name, ketywords, date last modified)

    Also set the subject, title and author properties.

    Usage:
      run DocProps algs.odp
      run DocProps story.doc 
*/


import com.sun.star.uno.*;

import it.naturtalent.libreoffice.utils.Info;
import it.naturtalent.libreoffice.utils.Lo;

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;


public class DocProps
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run DocProps <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    Info.printDocProperties(doc);

    Info.setDocProps(doc, "Example", "Examples", "Andrew Davison");
              // set subject, title, and author props
    Lo.save(doc);   // must save or the doc props are lost

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of DocProps class
