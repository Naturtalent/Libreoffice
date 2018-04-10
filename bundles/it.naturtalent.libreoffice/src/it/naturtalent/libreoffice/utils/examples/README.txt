
JLOP Chapter 1. Starting to Program with LibreOffice

From the website:

  Java LibreOffice Programming
  http://fivedots.coe.psu.ac.th/~ad/jlop

  Dr. Andrew Davison
  Dept. of Computer Engineering
  Prince of Songkla University
  Hat yai, Songkhla 90112, Thailand
  E-mail: ad@fivedots.coe.psu.ac.th


If you use this code, please mention my name, and include a link
to the website.

Thanks,
  Andrew

============================

This directory contains 7 Java files:
  * DispatchTest.java
  * DocConverter.java
  * DocInfo.java
  * DocMonitor.java
  * DocWindow.java
  * ExamineDoc.java
  * OfficeInfo.java



There are 5 batch files:
  * compile.bat
  * run.bat
     - make sure they refer to the correct locations for your
       Java, JNA, and my Utils classes.
       For details, read:
         "Installing the code for "Java LibreOffice Programming"
          http://fivedots.coe.psu.ac.th/~ad/jlop/install.html

  compile.bat and run.bat file both call:

  * lofind.bat
     - this tries to find LibreOffice (v4 or v5) on your machine,
       and checks its "bitness" with that of Java. They should
       both either be 32-bit or 64-bit. A warning is printed if 
       they are not.

  * loKill.bat        // kills the Office process
  * loList.bat        // lists the Office process (if it is running)


Extras: 
  * MRI-1.2.4.oxt     // the MRI inspection tool (see Section 12 of chapter)

  * algs.odp,          // assorted test files for the Java code
    points.odp, points.ppt
    skinner.jpg, skinner.png
    story.doc, story.pdf


----------------------------
Compilation:

> compile *.java
    // you must have LibreOffice, Java, JNA, and my Utils classes installed;
    // see above for details

----------------------------
Execution:

// You must have LibreOffice, Java, JNA, and my Utils classes installed;
// see above for details.

> run OfficeInfo


> run DispatchTest <fnm>
e.g.
> run DispatchTest algs.odp


> run DocConverter <fnm> <extension of saved doc>
e.g. 
> run DocConverter story.doc pdf


> run DocInfo <fnm>
e.g.
> run DocInfo story.doc


> run DocMonitor <fnm>
e.g.
> run DocMonitor algs.odp


> run DocWindow <fnm>
e.g.
> run DocWindow algs.odp


> run ExamineDoc <fnm>
e.g.
> run ExamineDoc story.doc
     // this example requires that MRI has been installed;
     // see section 12 of the chapter

----------------------------
Last updated: 15th August 2015
