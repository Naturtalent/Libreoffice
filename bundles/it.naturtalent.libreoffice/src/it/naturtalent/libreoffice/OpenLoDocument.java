package it.naturtalent.libreoffice;

import java.lang.reflect.InvocationTargetException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Status;
import org.eclipse.core.runtime.jobs.Job;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.swt.widgets.Display;

import com.sun.star.awt.XWindow;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.TerminationVetoException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XTerminateListener;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;

import it.naturtalent.libreoffice.utils.GUI;
import it.naturtalent.libreoffice.utils.Lo;

/**
 * 
 * 
 * @author Dieter Apel
 *
 */
public class OpenLoDocument
{

	private static Log log = LogFactory.getLog(OpenLoDocument.class);
	
	private static XComponentLoader officeDocumentLoader = null;
	
	// kill Watchdog
	private static boolean cancelWatchDog = false;
	
	
	/**
	 * Dokument in einem Job oeffnen
	 * 
	 * @param documentPath
	 */
	public static void loadLoDocument(final String documentPath)
	{				
		final Job j = new Job("Load Job") //$NON-NLS-1$
		{
			@Override
			protected IStatus run(final IProgressMonitor monitor)
			{
				try
				{					
					doLoadLoDocument(documentPath);
					// setDocumentProperties();
				} catch (Exception e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				return Status.OK_STATUS;
			}
		};

		j.schedule();
		runWatchdog();
	}	

	/**
	 * @param documentPath
	 */
	private static void doLoadLoDocument(final String documentPath)
	{
		// zur Faehigkeit ein Dokument zu laden braucht man einen XComponentLoader
		if (officeDocumentLoader == null)
		{
			// XComponentLoader ermitteln (ggf. LO starten)
			log.info("Libreoffice wurde gestartet");
			officeDocumentLoader = Lo.getOfficeLoader();
		}

		if (officeDocumentLoader != null)
		{					
			// Listener ueberwacht 'Libreoffice beenden'
			XDesktop xDesktop = Lo.getDesktop();
			xDesktop.addTerminateListener(new XTerminateListener()
			{
				public void queryTermination(EventObject e)
						throws TerminationVetoException
				{
					//System.out.println("TL: Starting Closing");
				}

				public void notifyTermination(EventObject e)
				{
					// Lo beendet - Loader ist nicht mehr gueltig
					officeDocumentLoader = null;
					log.info("Libreoffice wurde gestopped");
				}

				public void disposing(EventObject e)
				{
					//System.out.println("TL: Disposing");										
				}
			});
			
			// das Dokument wird geladen
			XComponent xDocument = Lo.openDoc(documentPath,officeDocumentLoader);
			if (xDocument != null)
			{
				log.info("Dokument: "+documentPath+" geöffnet");
				cancelWatchDog = true;
				
				// ContainerWindow (XWindow) wird sichtbar und erhaelt den
				// Focos
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
							log.info("Dokumentfenster geschlossen");
						}
					});
				}
			}
		}
		else log.info("kein officeDocumentLoader ermittelt");
	}
	

	
	/*
	 * 
	 */
	private static void runWatchdog()
	{
		ProgressMonitorDialog dialog = new ProgressMonitorDialog(Display.getDefault().getActiveShell());				
		dialog.open();
		try
		{
			dialog.run(true, true, new IRunnableWithProgress()
			{
				@Override
				public void run(IProgressMonitor monitor)
						throws InvocationTargetException,
						InterruptedException
				{
					monitor.beginTask("Libreoffice-Dokument wird geöffnet",IProgressMonitor.UNKNOWN);
					for (int i = 0;; ++i)
					{
						if (monitor.isCanceled())
						{
							throw new InterruptedException();
						}
						
						if (i == 50)
							break;
						if (cancelWatchDog)
							break;
						try
						{
							Thread.sleep(500);
						} catch (InterruptedException e)
						{
							throw new InterruptedException();
						}
					}
					monitor.done();							
				}
			});
		} catch (Exception e)
		{
			e.printStackTrace();
		}
	}

}
