package it.naturtalent.libreoffice.utils.examples;


// DocInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, March 2015

/* Print out information about the supplied file:
     * extension
     * office format represented by the extension
     * file type based on examining the file's metadata
     * properties associated with the file type
     * Office services that support this document

   Usage:
     compile DocInfo.java

     run DocInfo <fnm>
e.g.
     run DocInfo points.ppt > info.txt
     run DocInfo story.doc
     run DocInfo skinner.jpg

     run DocInfo "http://fivedots.coe.psu.ac.th/~ad/index.html"
*/


import com.sun.star.uno.*;

import it.naturtalent.libreoffice.utils.Info;
import it.naturtalent.libreoffice.utils.Lo;
import it.naturtalent.libreoffice.utils.Props;

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;



public class DocInfo
{

  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: DocInfo fnm");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();

    String ext = Info.getExt(args[0]);
    if (ext != null) {
      System.out.println("\nFile Extension: " + ext);
      System.out.println("Extension format: " + Lo.ext2Format(ext));
    }

    // get document type
    String docType = Info.getDocType(args[0]);
    if (docType != null) {
      System.out.println("Doc type: " + docType + "\n");
      Props.showDocTypeProps(docType);
    }


    // open the document so its services can be retrieved
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }


    System.out.println("\n------------ Services for this document: -----------");
    for(String service : Info.getServices(doc))
      System.out.println("  " + service);

    System.out.println("-----------");
    System.out.println("\n" + Lo.WRITER_SERVICE + " is supported: " +
                                  Info.isDocType(doc, Lo.WRITER_SERVICE) );


    System.out.println("\n------------ Available Services for this document: -----------");
    int count = 0;
    for(String service : Info.getAvailableServices(doc)) {
      System.out.println("  " + service);
      count++;
    }
    System.out.println("No. available services: " + count);


    System.out.println("\n----------------- Interfaces for this document: --------------");
    count = 0;
    for(String intfs : Info.getInterfaces(doc)) {
      System.out.println("  " + intfs);
      count++;
    }
    System.out.println("No. interfaces: " + count);


    String interfaceName = "com.sun.star.text.XTextDocument";
    System.out.println("\n------ Methods for interface " + interfaceName + ": ------");
    String[] methods = Info.getMethods(interfaceName);
    for(String methodName : methods)
      System.out.println("  " + methodName + "();");
    System.out.println("No. methods: " + methods.length);


    System.out.println("\n------------ Properties for this document: -----------------");
    count = 0;
    for(Property p : Props.getProperties(doc)) {
      System.out.println("  " + Props.showProperty(p));
      count++;
    }
    System.out.println("No. properties: " + count);

    System.out.println("\n-----------");

    // or in one line:
    // Props.showObjProps("Document", doc);

    String propName = "CharacterCount";
    System.out.println("Value of " + propName + ": " +
                                       Props.getProperty(doc, propName) );
    
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()


}  // end of DocInfo class

