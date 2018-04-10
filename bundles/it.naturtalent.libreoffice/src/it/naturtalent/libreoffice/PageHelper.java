package it.naturtalent.libreoffice;

/*************************************************************************
 *
 *  The Contents of this file are made available subject to the terms of
 *  the BSD license.
 *
 *  Copyright 2000, 2010 Oracle and/or its affiliates.
 *  All rights reserved.
 *
 *  Redistribution and use in source and binary forms, with or without
 *  modification, are permitted provided that the following conditions
 *  are met:
 *  1. Redistributions of source code must retain the above copyright
 *     notice, this list of conditions and the following disclaimer.
 *  2. Redistributions in binary form must reproduce the above copyright
 *     notice, this list of conditions and the following disclaimer in the
 *     documentation and/or other materials provided with the distribution.
 *  3. Neither the name of Sun Microsystems, Inc. nor the names of its
 *     contributors may be used to endorse or promote products derived
 *     from this software without specific prior written permission.
 *
 *  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
 *  "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 *  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
 *  FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
 *  COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 *  INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 *  BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS
 *  OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 *  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR
 *  TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
 *  USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 *************************************************************************/

// __________ Imports __________

import org.apache.commons.lang3.StringUtils;

import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XMasterPageTarget;
import com.sun.star.drawing.XMasterPagesSupplier;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.presentation.XHandoutMasterSupplier;
import com.sun.star.presentation.XPresentationPage;
import com.sun.star.uno.Any;
import com.sun.star.uno.UnoRuntime;


public class PageHelper
{
    // __________ static helper methods __________

    // __________ draw pages __________

	
	
    /** get the page count for standard pages
    */
    static public int getDrawPageCount( XComponent xComponent )
    {
        XDrawPagesSupplier xDrawPagesSupplier =
            UnoRuntime.queryInterface(
                    XDrawPagesSupplier.class, xComponent );
        XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();        
        return xDrawPages.getCount();
    }
    
    static public XDrawPage getDrawPageByName(XComponent xComponent, String pageName)
    {
        try
		{
			XDrawPagesSupplier xDrawPagesSupplier =
			    UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent );
			
			XDrawPage xDrawPage;
			int count = getDrawPageCount(xComponent);
			for(int i = 0;i < count;i++)
			{
					xDrawPage = getDrawPageByIndex(xComponent, i);				
					XPropertySet xPageProperties = UnoRuntime.queryInterface(
							XPropertySet.class, xDrawPage);
					
					String name = getPageName(xDrawPage);
					
					//String name = (String) xPageProperties.getPropertyValue("LinkDisplayName");
					if(StringUtils.equals(name, pageName))
						return getDrawPageByIndex(xComponent, i);
			}
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
        return null;
    }

    static public String getPageName(Object drawPage)
    {
    	if (drawPage instanceof XDrawPage)
		{
			XDrawPage xDrawPage = (XDrawPage) drawPage;
			return getPageName(xDrawPage);			
		}
    	
        return null;        
    }

    static public String getPageName(XDrawPage drawPage)
    {
    	 // each drawpage is supporting an XNamed interface
        XNamed xNamed = UnoRuntime.queryInterface(XNamed.class, drawPage);
        return xNamed.getName();        
    }
    
     static public void setPageName(XDrawPage drawPage, String pageName)
    {
    	 // each drawpage is supporting an XNamed interface
        XNamed xNamed = UnoRuntime.queryInterface(XNamed.class, drawPage);

        // beware, the page must have an unique name
        xNamed.setName(pageName);        
    }

     static public String getCurrentPage(XComponent xComponent)
     {
 		try
 		{
 			XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
 			XController xController = xModel.getCurrentController();
 			XPropertySet xPageProperties = UnoRuntime.queryInterface(
 					XPropertySet.class, xController);
 			
 			//Utils.printPropertyValues(xPageProperties);
 			
 			Any any = (Any) xPageProperties.getPropertyValue("CurrentPage");
 			XDrawPage xDrawPage = UnoRuntime.queryInterface(
 					XDrawPage.class, any);
 			
 			return getPageName(xDrawPage);		
 			
 		} catch (Exception e)
 		{
 			// TODO Auto-generated catch block
 			e.printStackTrace();
 		}
 		
 		return null;
 	}

    static public void setCurrentPage(XComponent xComponent, XDrawPage currentPage)
    {
		try
		{
			XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
			XController xController = xModel.getCurrentController();
			XPropertySet xPageProperties = UnoRuntime.queryInterface(
					XPropertySet.class, xController);			
			xPageProperties.setPropertyValue("CurrentPage", currentPage);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

    /** get draw page by index
    */
    static public XDrawPage getDrawPageByIndex( XComponent xComponent, int nIndex )
        throws com.sun.star.lang.IndexOutOfBoundsException,
            com.sun.star.lang.WrappedTargetException
    {
        XDrawPagesSupplier xDrawPagesSupplier =
            UnoRuntime.queryInterface(
                    XDrawPagesSupplier.class, xComponent );
        XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
        return UnoRuntime.queryInterface(XDrawPage.class, xDrawPages.getByIndex( nIndex ));
    }

    /** creates and inserts a draw page into the giving position,
        the method returns the new created page
    */
    static public XDrawPage insertNewDrawPageByIndex( XComponent xComponent, int nIndex )
        throws Exception
    {
        XDrawPagesSupplier xDrawPagesSupplier =
            UnoRuntime.queryInterface(
                    XDrawPagesSupplier.class, xComponent );
        XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
        return xDrawPages.insertNewByIndex( nIndex );
    }

    /** removes the given page
    */
    static public void removeDrawPage( XComponent xComponent, XDrawPage xDrawPage )
    {
        XDrawPagesSupplier xDrawPagesSupplier =
            UnoRuntime.queryInterface(
                    XDrawPagesSupplier.class, xComponent );
        XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
        xDrawPages.remove( xDrawPage );
    }

    /** get size of the given page
    */
    static public Size getPageSize( XDrawPage xDrawPage )
        throws com.sun.star.beans.UnknownPropertyException,
            com.sun.star.lang.WrappedTargetException
    {
        XPropertySet xPageProperties = UnoRuntime.queryInterface( XPropertySet.class, xDrawPage );
        return new Size(
            ((Integer)xPageProperties.getPropertyValue( "Width" )).intValue(),
            ((Integer)xPageProperties.getPropertyValue( "Height" )).intValue() );
    }

    // __________ master pages __________

    /** get the page count for master pages
    */
    static public int getMasterPageCount( XComponent xComponent )
    {
        XMasterPagesSupplier xMasterPagesSupplier =
            UnoRuntime.queryInterface(
                    XMasterPagesSupplier.class, xComponent );
        XDrawPages xDrawPages = xMasterPagesSupplier.getMasterPages();
        return xDrawPages.getCount();
    }

    /** get master page by index
    */
    static public XDrawPage getMasterPageByIndex( XComponent xComponent, int nIndex )
        throws com.sun.star.lang.IndexOutOfBoundsException,
            com.sun.star.lang.WrappedTargetException
    {
        XMasterPagesSupplier xMasterPagesSupplier =
            UnoRuntime.queryInterface(
                    XMasterPagesSupplier.class, xComponent );
        XDrawPages xDrawPages = xMasterPagesSupplier.getMasterPages();
        return UnoRuntime.queryInterface(XDrawPage.class, xDrawPages.getByIndex( nIndex ));
    }

    /** creates and inserts a new master page into the giving position,
        the method returns the new created page
    */
    static public XDrawPage insertNewMasterPageByIndex( XComponent xComponent, int nIndex )
    {
        XMasterPagesSupplier xMasterPagesSupplier =
            UnoRuntime.queryInterface(
                    XMasterPagesSupplier.class, xComponent );
        XDrawPages xDrawPages = xMasterPagesSupplier.getMasterPages();
        return xDrawPages.insertNewByIndex( nIndex );
    }

    /** removes the given page
    */
    static public void removeMasterPage( XComponent xComponent, XDrawPage xDrawPage )
    {
        XMasterPagesSupplier xMasterPagesSupplier =
            UnoRuntime.queryInterface(
                    XMasterPagesSupplier.class, xComponent );
        XDrawPages xDrawPages = xMasterPagesSupplier.getMasterPages();
        xDrawPages.remove( xDrawPage );
    }

    /** return the corresponding masterpage for the giving drawpage
    */
    static public XDrawPage getMasterPage( XDrawPage xDrawPage )
    {
        XMasterPageTarget xMasterPageTarget =
            UnoRuntime.queryInterface(
                    XMasterPageTarget.class, xDrawPage );
        return xMasterPageTarget.getMasterPage();
    }

    /** sets given masterpage at the drawpage
    */
    static public void setMasterPage( XDrawPage xDrawPage, XDrawPage xMasterPage )
    {
        XMasterPageTarget xMasterPageTarget =
            UnoRuntime.queryInterface(
                    XMasterPageTarget.class, xDrawPage );
        xMasterPageTarget.setMasterPage( xMasterPage );
    }

    // __________ presentation pages __________

    /** test if a Presentation Document is supported.
        This is important, because only presentation documents
        have notes and handout pages
    */
    static public boolean isImpressDocument( XComponent xComponent )
    {
        XServiceInfo xInfo = UnoRuntime.queryInterface(
                XServiceInfo.class, xComponent );
        return xInfo.supportsService( "com.sun.star.presentation.PresentationDocument" );
    }

    /** in impress documents each normal draw page has a corresponding notes page
    */
    static public XDrawPage getNotesPage( XDrawPage xDrawPage )
    {
        XPresentationPage aPresentationPage =
            UnoRuntime.queryInterface(
                    XPresentationPage.class, xDrawPage );
        return aPresentationPage.getNotesPage();
    }

    /** in impress each documents has one handout page
    */
    static public XDrawPage getHandoutMasterPage( XComponent xComponent )
    {
        XHandoutMasterSupplier aHandoutMasterSupplier =
            UnoRuntime.queryInterface(
                    XHandoutMasterSupplier.class, xComponent );
        return aHandoutMasterSupplier.getHandoutMasterPage();
    }
}
