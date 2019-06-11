package it.naturtalent.libreoffice.text;

import java.io.File;

import org.apache.commons.lang3.SystemUtils;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Status;
import org.eclipse.core.runtime.jobs.Job;
import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.internal.workbench.E4Workbench;
import org.eclipse.e4.ui.model.application.MApplication;
import org.eclipse.e4.ui.workbench.IWorkbench;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.swt.widgets.Display;

import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.FrameActionEvent;
import com.sun.star.frame.TerminationVetoException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XFrameActionListener;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XTerminateListener;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XSelectionChangeListener;
import com.sun.star.view.XSelectionSupplier;

import it.naturtalent.libreoffice.Activator;
import it.naturtalent.libreoffice.Bootstrap;
import it.naturtalent.libreoffice.draw.TerminateListener;
import it.naturtalent.libreoffice.utils.GUI;
import it.naturtalent.libreoffice.utils.JNAUtils;
import it.naturtalent.libreoffice.utils.Lo;
import it.naturtalent.libreoffice.utils.examples.DocMonitor;

public class TextDocument
{
	
	public static final String TEXTDOCUMENT_EVENT_DOCUMENT_OPEN = "textDocumentOpen"; //$NON-NLS-N$
	public static final String TEXTDOCUMENT_EVENT_DOCUMENT_CLOSE = "textDocumentClose"; //$NON-NLS-N$
	
	private String documentPath;	
	private XComponentContext xContext;
	private XComponent xComponent;
	private IEventBroker eventBroker;
	private XDesktop xDesktop;
	
	private static final String libPath = "/usr/lib/libreoffice/program";
	
	// echo -e PATH=\"/usr/lib/libreoffice/program:\$PATH\" >> $HOME/.profile
	
	public void loadPage(final String documentPath)
	{
		
		//it.naturtalent.libreoffice.utils.JNAUtils.killOffice();
		
		// Libreoffice-Dokumentenlader abrufen
		/*
		XComponentLoader xComponentLoader = Lo.getOfficeLoader();		
		Lo.delay(5000);
		Lo.closeOffice();
		if(xComponentLoader != null)
			return;
			*/
		
		// einen Documentloader anfragen (ggf. Libreoffice starten) 
		// (XComponentContext, XDesktop, XMultiComponentFactory sind statisch verfuegbar)
		XComponentLoader xComponentLoader = Lo.getOfficeLoader();		
		if(xComponentLoader != null)
		{
			// Listener ueberwacht 'Libreoffice beenden'
			XDesktop xDesktop = Lo.getDesktop();
			xDesktop.addTerminateListener(new XTerminateListener()
			{
				public void queryTermination(EventObject e)
						throws TerminationVetoException
				{
					System.out.println("TL: Starting Closing");
				}

				public void notifyTermination(EventObject e)
				{
					System.out.println("TL: Finished Closing");
				}

				public void disposing(EventObject e)
				{
					System.out.println("TL: Disposing");										
				}
			});

			// das Dokument wird geladen
			XComponent xDocument = Lo.openDoc(documentPath,xComponentLoader);
			if (xDocument != null)
			{		
				// ContainerWindow (XWindow) wird sichtbar und erhaelt den Focos
				GUI.setVisible(xDocument, true);
				
				// Listener ueberwacht das Schliessen des Dokuments
				XWindow xWindow = GUI.getFrame(xDocument).getContainerWindow();				
				if (xWindow != null)
				{
					xWindow.addEventListener(new XEventListener()
					{
						@Override
						public void disposing(EventObject arg0)
						{
							System.out.println("Fenster geschlossen");

						}
					});
				}
			}
			else
			{
				Lo.delay(5000);
				Lo.closeOffice();
			}
		}
		
		
		
		/*
		MApplication currentApplication = E4Workbench.getServiceContext().get(IWorkbench.class).getApplication();
		eventBroker = currentApplication.getContext().get(IEventBroker.class);
		
		Job j = new Job("Load Job") //$NON-NLS-1$
		{
			@Override
			protected IStatus run(final IProgressMonitor monitor)
			{
				try
				{
					loadDocumentLo(documentPath);
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
		*/		
	}
	
	/*
	 * 
	 * 
	 * 
	 * 
	 */

	private void loadDocumentLo(String documentPath) throws Exception
	{
		/*
		XComponentContext context = Activator.getOfficeContext();
		if(context != null)
		{
			XComponentLoader xComponentLoader = UnoRuntime.queryInterface(XComponentLoader.class, Lo.getDesktop());
			
			XComponent xComponent = Lo.openDoc(documentPath, xComponentLoader);
			
		}
		*/
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
