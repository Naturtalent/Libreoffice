package it.naturtalent.libreoffice;

import java.io.File;
import java.io.IOException;

import org.apache.commons.lang3.ArrayUtils;

import com.sun.star.beans.PropertyValue;
import com.sun.star.document.XFilter;
import com.sun.star.document.XImporter;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import it.naturtalent.libreoffice.draw.DrawDocument;

public class ImportFilter
{
	private DrawDocument drawDocument;

	public ImportFilter(DrawDocument drawDocument)
	{
		super();
		this.drawDocument = drawDocument;
	}

	
	public void doImport(File sourceFile)
	{
		XComponentContext xContext = drawDocument.getxContext();
		XMultiComponentFactory xComponentFactory = xContext.getServiceManager();
		
        // create filter
		try
		{		
			XMultiComponentFactory xmulticomponentfactory = xContext.getServiceManager();
			//XMultiComponentFactory xmulticomponentfactory = (XMultiComponentFactory) UnoRuntime.queryInterface( XMultiComponentFactory.class, xLocalContext );
			
			String [] names = xmulticomponentfactory.getAvailableServiceNames();
			
			String filter = "com.sun.star.document.ImportFilter";
			if (ArrayUtils.contains(names, filter))
			{
				Object importFilter = xmulticomponentfactory
						.createInstanceWithContext(filter,xContext);
				XFilter xfilter = (XFilter) UnoRuntime.queryInterface(
						XFilter.class, importFilter);
				
				Object extTypeDet = xmulticomponentfactory
						.createInstanceWithContext("com.sun.star.document.ExtendedTypeDetection",xContext);
				
								
				XImporter ximporter = (XImporter) UnoRuntime.queryInterface(
						XImporter.class, importFilter);
				ximporter.setTargetDocument(drawDocument.getxComponent());
				
				 PropertyValue[] propertyvalue = new PropertyValue[2];
                 propertyvalue[ 0 ] = new PropertyValue();
                 propertyvalue[ 0 ].Name = "FileName";

                 propertyvalue[ 1 ] = new PropertyValue();
                 propertyvalue[ 1 ].Name = "PagePos";

                 StringBuffer sSaveUrl = new StringBuffer("file:///");
                 try
				{
					sSaveUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
	                propertyvalue[ 0 ].Value = sSaveUrl.toString();
			        propertyvalue[ 1 ].Value = new Integer( 1 );
				  	xfilter.filter( propertyvalue );

				} catch (IOException e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}

		} catch (com.sun.star.uno.Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}


	
}
