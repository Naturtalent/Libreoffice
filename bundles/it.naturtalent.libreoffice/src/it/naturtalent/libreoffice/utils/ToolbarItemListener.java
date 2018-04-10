package it.naturtalent.libreoffice.utils;


// ToolbarItemListener.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, August 2016, 

/* Interface used by objects passed to an ItemInterceptor
   object. See ItemInterceptor.java in this folder.
*/

import com.sun.star.beans.*;
import com.sun.star.util.*;


public interface ToolbarItemListener 
{
  void clicked(String itemName, URL cmdURL, PropertyValue[] props);

}
