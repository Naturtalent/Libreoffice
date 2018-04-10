package it.naturtalent.libreoffice.utils.examples;


// OfficeInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, March 2015

/* Report OS info (using Java), Office configuration details, and 
   Office paths. Optionally list all of Office's service names (BIG),
   and filter names.

   Usage:
     > compile OfficeInfo.java
     > run OfficeInfo
*/


import com.sun.star.uno.*;

import it.naturtalent.libreoffice.utils.Info;
import it.naturtalent.libreoffice.utils.Lo;

import com.sun.star.lang.*;
import com.sun.star.frame.*;


public class OfficeInfo
{

  public static void main(String[] args)
  {
    System.out.println("OS Name: " + System.getProperty("os.name"));
    System.out.println("OS Version: " + System.getProperty("os.version"));
    System.out.println("OS Architecture: " + System.getProperty("os.arch"));

    XComponentLoader loader = Lo.loadOffice();
                              // Lo.loadSocketOffice();
                             
                              
    System.out.println("\nOffice name: " + Info.getConfig("ooName"));

    System.out.println("\nOffice version (long): " + 
                             Info.getConfig("ooSetupVersionAboutBox"));
    System.out.println("Office version (short): " + Info.getConfig("ooSetupVersion"));

    System.out.println("\nOffice language location: " + Info.getConfig("ooLocale"));
    System.out.println("System language location: \"" + Info.getConfig("ooSetupSystemLocale") + "\"");


    System.out.println("\nWorking Dir: " + Info.getPaths("Work"));
    System.out.println("\nOffice Dir: " + Info.getOfficeDir());

    System.out.println("\nAddin Dir: " + Info.getPaths("Addin"));
    System.out.println("\nFilters Dir: " + Info.getPaths("Filter"));
    System.out.println("\nTemplates Dirs: " + Info.getPaths("Template"));
    System.out.println("\nGallery Dir: " + Info.getPaths("Gallery"));
      // see https://wiki.openoffice.org/w/index.php?title=Documentation/DevGuide/OfficeDev/Path_Settings


/*
    System.out.println("\n---------- Services for Office: -----------------");
    for(String service : Info.getServiceNames())
      System.out.println("  " + service);
    System.out.println("-----------");
    System.out.println("No. of services: " + Info.getServiceNames().length);
              // 1000+ printed!!
*/

/*
    System.out.println("\n-----------");
    System.out.println("Services offered by the service " + Lo.WRITER_SERVICE + ":");
    String[] nms = Info.getServiceNames(Lo.WRITER_SERVICE);
    for(String service : nms)
      System.out.println("  " + service);
*/

/*
    System.out.println("\n---------- File Filter Names for Office: -----------------");
    for(String nms : Info.getFilterNames())
      System.out.println("  " + nms);
    System.out.println("-----------");
    System.out.println("No. of filter names: " + Info.getFilterNames().length);
              //  about 250 printed!!
*/

    System.out.println();
    Lo.closeOffice();
  } // end of main()


}  // end of OfficeInfo class

