package it.naturtalent.libreoffice.utils.examples;


// DocConverter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, March 2015

/*  Convert a document to another format. The new file uses
    the name of the old file, but with the new extension.

    Usage: run DocConverter <fnm> <extension of saved doc>

    e.g. a 'text' document can be converted to doc, docx, rtf, odt, pdf, txt;
          run DocConverter story.doc pdf

          run DocConverter "http://fivedots.coe.psu.ac.th/~ad/index.html" pdf


    e.g. a 'presentation' document can be converted to ppt pptx, odp
               run DocConverter points.ppt odp
       // conversion of ppt-->odp can take a few minutes for 30+ slides!

    e.g. a 'draw' document can be converted to png, jpg, odg
               run DocConverter skinner.jpg png

    The limitation is that Lo.saveDoc() uses a hardwired mapping of 
    save extension to office filter name, inside Lo.ext2Format().

    If the extension isn't listed, then the export will fail, 
    issuing an com.sun.star.task.ErrorCodeIOException
*/

import com.sun.star.uno.*;

import it.naturtalent.libreoffice.utils.Info;
import it.naturtalent.libreoffice.utils.Lo;

import com.sun.star.lang.*;
import com.sun.star.frame.*;


public class DocConverter 
{

  public static void main(String args[]) 
  {
    if (args.length != 2) {
      System.out.println("Usage: DocConverter fnm extension");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();

    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    String name = Info.getName(args[0]);
    Lo.saveDoc(doc, name + "." + args[1]);
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()

}  // end of DocConverter class
