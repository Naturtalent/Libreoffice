package it.naturtalent.libreoffice.text;

import it.naturtalent.libreoffice.Bootstrap;
import it.naturtalent.libreoffice.DrawDocumentEvent;
import it.naturtalent.libreoffice.draw.TerminateListener;

import java.io.File;

import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Status;
import org.eclipse.core.runtime.jobs.Job;
import org.eclipse.e4.core.services.events.IEventBroker;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XSelectionChangeListener;
import com.sun.star.view.XSelectionSupplier;

public class TextDocument
{
	
	public static final String TEXTDOCUMENT_EVENT_DOCUMENT_OPEN = "textDocumentOpen"; //$NON-NLS-N$
	public static final String TEXTDOCUMENT_EVENT_DOCUMENT_CLOSE = "textDocumentClose"; //$NON-NLS-N$
	
	private String documentPath;	
	private XComponentContext xContext;
	private XComponent xComponent;
	private IEventBroker eventBroker;
	private XDesktop xDesktop;
	
	public void loadPage(final String documentPath)
	{
		Job j = new Job("Load Job") //$NON-NLS-1$
		{
			@Override
			protected IStatus run(final IProgressMonitor monitor)
			{
				try
				{
					loadDocument(documentPath);
					//setDocumentProperties();
				}
				catch (Exception e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
				return Status.OK_STATUS;
			}
		};
		j.schedule();		
	}

	private void loadDocument(String documentPath) throws Exception
	{
		this.documentPath = documentPath;
		
		File sourceFile = new java.io.File(documentPath);
		StringBuffer sTemplateFileUrl = new StringBuffer("file:///");
		sTemplateFileUrl.append(sourceFile.getCanonicalPath()
				.replace('\\', '/'));
	
		xContext = Bootstrap.bootstrap();
		if (xContext != null)
		{
			XMultiComponentFactory xMCF = xContext.getServiceManager();
	
			// retrieve the Desktop object, we need its XComponentLoader
			Object desktop = xMCF.createInstanceWithContext(
					"com.sun.star.frame.Desktop", xContext);
	
			XComponentLoader xComponentLoader = UnoRuntime.queryInterface(
					XComponentLoader.class, desktop);
	
			// load
			PropertyValue[] loadProps = new PropertyValue[0];
			xComponent = xComponentLoader.loadComponentFromURL(
					sTemplateFileUrl.toString(), "_blank", 0, loadProps);
			
			xComponent.addEventListener(new XEventListener()
			{				
				@Override
				public void disposing(EventObject arg0)
				{
					if(eventBroker != null)
						eventBroker.post(TEXTDOCUMENT_EVENT_DOCUMENT_CLOSE, this);					
				}
			});
			
			// TerminateListener installieren
			xDesktop = UnoRuntime.queryInterface(XDesktop.class,desktop);
			TerminateListener terminateListener = new TerminateListener();
			xDesktop.addTerminateListener(terminateListener);
			terminateListener.setEventBroker(eventBroker);
			
			if(eventBroker != null)
				eventBroker.post(TEXTDOCUMENT_EVENT_DOCUMENT_OPEN, this);
			
			XModel xModel = UnoRuntime.queryInterface(XModel.class,xComponent);
			XController xController = xModel.getCurrentController();
			
			XSelectionSupplier selectionSupplier = UnoRuntime.queryInterface(
					XSelectionSupplier.class, xController);
			
			selectionSupplier.addSelectionChangeListener(new XSelectionChangeListener()
			{
				
				@Override
				public void disposing(EventObject arg0)
				{
					// TODO Auto-generated method stub
					
				}
				
				@Override
				public void selectionChanged(EventObject arg0)
				{
					/*
					Object obj = arg0.Source;
					if (obj instanceof XPropertySet)
					{
						XPropertySet selectedPropertySet = (XPropertySet) obj;
						Utils.printPropertyValues(selectedPropertySet);
					}
					
					System.out.println("Selection: "+obj);
					*/
				
					
				}
			});
			
			// das geoeffnete Dokument mit Listener als Key speichern
			//openDocumentsMap.put(terminateListener, documentPath);	
		}
	}

}
