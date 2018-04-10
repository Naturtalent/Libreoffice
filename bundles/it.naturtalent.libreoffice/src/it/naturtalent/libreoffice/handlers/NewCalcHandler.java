 
package it.naturtalent.libreoffice.handlers;

import java.io.File;
import java.io.IOException;
import java.net.URL;

import javax.inject.Named;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.eclipse.core.resources.IContainer;
import org.eclipse.core.resources.IResource;
import org.eclipse.core.runtime.FileLocator;
import org.eclipse.core.runtime.IPath;
import org.eclipse.core.runtime.Path;
import org.eclipse.e4.core.di.annotations.CanExecute;
import org.eclipse.e4.core.di.annotations.Execute;
import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.internal.workbench.E4Workbench;
import org.eclipse.e4.ui.model.application.MApplication;
import org.eclipse.e4.ui.model.application.ui.basic.MPart;
import org.eclipse.e4.ui.services.IServiceConstants;
import org.eclipse.e4.ui.workbench.IWorkbench;
import org.eclipse.e4.ui.workbench.modeling.EPartService;
import org.eclipse.e4.ui.workbench.modeling.ESelectionService;
import org.eclipse.swt.widgets.Shell;
import org.osgi.framework.Bundle;
import org.osgi.framework.BundleContext;
import org.osgi.framework.FrameworkUtil;

import it.naturtalent.e4.project.ui.navigator.ResourceNavigator;
import it.naturtalent.e4.project.ui.utils.RefreshResource;
import it.naturtalent.libreoffice.Activator;
import it.naturtalent.libreoffice.calc.CalcDocument;
import it.naturtalent.libreoffice.draw.DrawDocument;


/**
 * Ein neues DrawFile erzeugen.
 * 
 * @author dieter
 *
 */
public class NewCalcHandler
{
	// Basisname der neuen Tabelle (Zieldatei)
	public static final String CALC_FILENAME = "calc.ods"; //$NON-NLS-1$
	
	// Template der neuen Tabelle benutze Vorlage (Quelle)  
	private static final String CALC_TEMPLATE = "/templates/calc.ods"; //$NON-NLS-1$
	
	@Execute
	public void execute(IEventBroker eventBroker, @Named(IServiceConstants.ACTIVE_SHELL) Shell shell)
	{
		MApplication currentApplication = E4Workbench.getServiceContext().get(IWorkbench.class).getApplication();
		EPartService partService = currentApplication.getContext().get(EPartService.class);
		MPart part = partService.findPart(ResourceNavigator.RESOURCE_NAVIGATOR_ID);
							
		ESelectionService selectionService = part.getContext().get(ESelectionService.class);
		Object selObject = selectionService.getSelection();
		if (selObject instanceof IResource)
		{
			IResource iResource = (IResource) selObject;			
			if(iResource.getType() == IResource.FILE)
				iResource = iResource.getParent();
			
			// neue Datei erzeugen
			String newName = getAutoFileName((IContainer) iResource, CALC_FILENAME);				
			IPath path = iResource.getLocation().append(newName);
			createDrawFile(path.toFile());		

			// Refresh IContainer (neue Datei im ResourceNavigator anzeigen)
			RefreshResource refreshResource = new RefreshResource();
			refreshResource.refresh(shell, iResource);
			
			// Zeichnung in LibreOffice anzeigen
			new CalcDocument().loadPage(path.toString());
		}
	}
	
	@CanExecute
	public boolean canExecute()
	{
		MApplication currentApplication = E4Workbench.getServiceContext().get(IWorkbench.class).getApplication();
		EPartService partService = currentApplication.getContext().get(EPartService.class);
		MPart part = partService.findPart(ResourceNavigator.RESOURCE_NAVIGATOR_ID);
							
		ESelectionService selectionService = part.getContext().get(ESelectionService.class);
		Object selObject = selectionService.getSelection();
		return (selObject instanceof IResource);
	}
	
	/*
	 * Kopiert die DrawTemplateDatei in das Zielverzeichnis.
	 */
	public static void createDrawFile(File destDrawFile)
	{
		Bundle bundle = FrameworkUtil.getBundle(Activator.class);
		BundleContext bundleContext = bundle.getBundleContext();
		URL urlTemplate = FileLocator.find(bundleContext.getBundle(),new Path(CALC_TEMPLATE), null);
		try
		{
			urlTemplate = FileLocator.resolve(urlTemplate);
			try
			{				
				FileUtils.copyURLToFile(urlTemplate, destDrawFile);				
			} catch (IOException e)
			{							
				//log.error(Messages.DesignUtils_ErrorCreateDrawFile);
				e.printStackTrace();
			}
			
		} catch (IOException e1)
		{						
			//log.error(Messages.DesignUtils_ErrorCreateDrawFile);
			e1.printStackTrace();
		}
	}
	
	/*
	 * 
	 */
	private String getAutoFileName(IContainer dir, String originalFileName)
	{
		String autoFileName;

		if (dir == null)
			return "";

		int counter = 1;
		while (true)
		{
			if (counter > 1)
			{
				autoFileName = FilenameUtils.getBaseName(originalFileName)
						+ new Integer(counter) + "."
						+ FilenameUtils.getExtension(originalFileName);
			}
			else
			{
				autoFileName = originalFileName;
			}

			IResource res = dir.findMember(autoFileName);
			if (res == null)
				return autoFileName;

			counter++;
		}
	}

	
	

}