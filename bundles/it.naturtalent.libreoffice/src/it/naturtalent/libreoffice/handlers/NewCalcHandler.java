 
package it.naturtalent.libreoffice.handlers;

import javax.inject.Inject;
import javax.inject.Named;

import org.eclipse.core.resources.IResource;
import org.eclipse.e4.core.di.annotations.CanExecute;
import org.eclipse.e4.core.di.annotations.Execute;
import org.eclipse.e4.core.di.annotations.Optional;
import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.services.IServiceConstants;
import org.eclipse.e4.ui.workbench.modeling.ESelectionService;
import org.eclipse.swt.widgets.Shell;

import it.naturtalent.e4.project.ui.navigator.ResourceNavigator;
import it.naturtalent.libreoffice.calc.CalcDocument;


/**
 * Ein neues CalcFile erzeugen.
 * 
 * @author dieter
 *
 */
public class NewCalcHandler extends NewTextHandler
{
	@Inject @Optional private ESelectionService selectionService;
	
	// Basisname der neuen Tabelle (Zieldatei)
	public static final String CALC_FILENAME = "calc.ods"; //$NON-NLS-1$
	
	@Execute
	public void execute(IEventBroker eventBroker, @Named(IServiceConstants.ACTIVE_SHELL) Shell shell)
	{
		Object selObject = selectionService.getSelection(ResourceNavigator.RESOURCE_NAVIGATOR_ID);
		if (selObject instanceof IResource)
		{	
			originalBaseFileName = CALC_FILENAME;
			templateFileName = TEMPLATE + CALC_FILENAME;
			
			IResource iResource = (IResource) selObject;			
			if(iResource.getType() == IResource.FILE)
				iResource = iResource.getParent();
			doExecute(eventBroker, iResource, shell);
		}	
	}
	
	@Override
	protected void loadDocument(String filePath)
	{
		new CalcDocument().loadPage(filePath);
	}

	@CanExecute
	public boolean canExecute()
	{				
		return (selectionService.getSelection(
				ResourceNavigator.RESOURCE_NAVIGATOR_ID) instanceof IResource);
	}
	


}