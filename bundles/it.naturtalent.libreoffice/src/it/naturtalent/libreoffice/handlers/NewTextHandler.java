 
package it.naturtalent.libreoffice.handlers;

import java.io.File;
import java.io.IOException;
import java.net.URL;

import javax.inject.Inject;
import javax.inject.Named;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.eclipse.core.resources.IContainer;
import org.eclipse.core.resources.IFile;
import org.eclipse.core.resources.IResource;
import org.eclipse.core.resources.IWorkspace;
import org.eclipse.core.resources.ResourcesPlugin;
import org.eclipse.core.runtime.FileLocator;
import org.eclipse.core.runtime.IPath;
import org.eclipse.core.runtime.Path;
import org.eclipse.e4.core.di.annotations.CanExecute;
import org.eclipse.e4.core.di.annotations.Execute;
import org.eclipse.e4.core.di.annotations.Optional;
import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.services.IServiceConstants;
import org.eclipse.e4.ui.workbench.modeling.ESelectionService;
import org.eclipse.swt.widgets.Shell;
import org.osgi.framework.Bundle;
import org.osgi.framework.BundleContext;
import org.osgi.framework.FrameworkUtil;

import it.naturtalent.e4.project.IResourceNavigator;
import it.naturtalent.e4.project.ui.navigator.ResourceNavigator;
import it.naturtalent.e4.project.ui.utils.RefreshResource;
import it.naturtalent.libreoffice.Activator;
import it.naturtalent.libreoffice.text.TextDocument;

public class NewTextHandler
{
	
	@Inject @Optional private ESelectionService selectionService;
	
	// Basisname der neuen Zeichnung (Zieldatei)
	public static final String TEXT_FILENAME = "text.odt"; //$NON-NLS-1$

	// Template der neuen Textdatei (Quelle)  
	protected static final String TEMPLATE = "/templates/"; //$NON-NLS-1$

	protected String originalBaseFileName = TEXT_FILENAME;
	
	protected String templateFileName = TEMPLATE+TEXT_FILENAME;
	
		
	@Execute
	public void execute(IEventBroker eventBroker,  @Named(IServiceConstants.ACTIVE_SHELL) Shell shell)
	{
		Object selObject = selectionService.getSelection(ResourceNavigator.RESOURCE_NAVIGATOR_ID);
		if (selObject instanceof IResource)
		{			
			IResource iResource = (IResource) selObject;			
			if(iResource.getType() == IResource.FILE)
				iResource = iResource.getParent();
			doExecute(eventBroker, iResource, shell);
		}
	}
	
	protected void doExecute(IEventBroker eventBroker, IResource iResource, Shell shell)
	{
		// neue Datei erzeugen
		String newName = getAutoFileName((IContainer) iResource);				
		IPath path = iResource.getLocation().append(newName);
		createFile(path.toFile());		
		
		// Refresh IContainer (neue Datei im ResourceNavigator anzeigen)
		RefreshResource refreshResource = new RefreshResource();
		refreshResource.refresh(shell, iResource);
		
		IWorkspace workspace= ResourcesPlugin.getWorkspace();  
		IPath location= Path.fromOSString(path.toFile().getAbsolutePath()); 
		IFile ifile= workspace.getRoot().getFileForLocation(location);
		eventBroker.post(IResourceNavigator.NAVIGATOR_EVENT_SELECT_REQUEST, ifile);
		
		// Zeichnung in LibreOffice anzeigen
		loadDocument(path.toString());

	}
	
	protected void loadDocument(String filePath)
	{
		new TextDocument().loadPage(filePath);
	}
	


	@CanExecute
	public boolean canExecute()
	{
		return (selectionService.getSelection(
				ResourceNavigator.RESOURCE_NAVIGATOR_ID) instanceof IResource);
	}
	
	/*
	 * Kopiert die DrawTemplateDatei in das Zielverzeichnis.
	 */
	public void createFile(File destFile)
	{
		Bundle bundle = FrameworkUtil.getBundle(Activator.class);
		BundleContext bundleContext = bundle.getBundleContext();
		URL urlTemplate = FileLocator.find(bundleContext.getBundle(),new Path(templateFileName), null);
		try
		{
			urlTemplate = FileLocator.resolve(urlTemplate);
			try
			{				
				FileUtils.copyURLToFile(urlTemplate, destFile);				
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
	protected String getAutoFileName(IContainer dir)
	{
		String autoFileName;

		if (dir == null)
			return "";

		int counter = 1;
		while (true)
		{
			if (counter > 1)
			{
				autoFileName = FilenameUtils.getBaseName(originalBaseFileName)
						+ new Integer(counter) + "."
						+ FilenameUtils.getExtension(originalBaseFileName);
			}
			else
			{
				autoFileName = originalBaseFileName;
			}

			IResource res = dir.findMember(autoFileName);
			if (res == null)
				return autoFileName;

			counter++;
		}
	}

		
}