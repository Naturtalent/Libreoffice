 
package it.naturtalent.libreoffice;

import java.util.List;

import javax.inject.Inject;

import org.apache.commons.lang3.StringUtils;
import org.eclipse.e4.core.di.annotations.Optional;
import org.eclipse.e4.core.di.extensions.EventTopic;
import org.eclipse.e4.ui.model.application.MApplication;
import org.eclipse.e4.ui.model.application.commands.MCommand;
import org.eclipse.e4.ui.workbench.UIEvents;
import org.osgi.service.event.Event;

import it.naturtalent.application.services.IOpenWithEditorAdapterRepository;
import it.naturtalent.e4.project.ui.DynamicNewMenu;

public class LibreOfficeAddon
{

	private @Inject @Optional MApplication application;
	
	@Inject @Optional private IOpenWithEditorAdapterRepository openwithAdapterRepository;
	
	
	// Dynamic New
	public static final String NEW_DRAW_MENUE_ID = "it.naturtalent.e4.libreoffice.menue.DrawDocument"; //$NON-NLS-1$
	public static final String NEW_DRAW_LABEL = "Zeichnung"; //$NON-NLS-1$
	public static final String NEW_DRAW_COMMAND_ID = "it.naturtalent.libreoffice.command.newdrawdocument"; //$NON-NLS-1$

	public static final String NEW_CALC_MENUE_ID = "it.naturtalent.e4.libreoffice.menue.CalcDocument"; //$NON-NLS-1$
	public static final String NEW_CALC_LABEL = "Kalkulation"; //$NON-NLS-1$
	public static final String NEW_CALC_COMMAND_ID = "it.naturtalent.libreoffice.command.newcalcdocument"; //$NON-NLS-1$

	public static final String NEW_TEXT_MENUE_ID = "it.naturtalent.e4.libreoffice.menue.TextDocument"; //$NON-NLS-1$
	public static final String NEW_TEXT_LABEL = "Text"; //$NON-NLS-1$
	public static final String NEW_TEXT_COMMAND_ID = "it.naturtalent.libreoffice.command.newtextdocument"; //$NON-NLS-1$

	// Position im KontextMenue
	private static final int DRAW_MENUE_POSITION = 3;
	
	@Inject
	@Optional
	public void applicationStarted(@EventTopic(UIEvents.UILifeCycle.APP_STARTUP_COMPLETE) Event event)
	{
		String label;
		DynamicNewMenu newMenu = new DynamicNewMenu();
		
		// Adapter zum Oeffnen von Dateien mit Libreoffice
		openwithAdapterRepository.getOpenWithAdapters().add(new LibreofficeOpenWithAdapter());
		//openwithAdapterRepository.getOpenWithAdapters().add(new OpenLoDocumentAdapter());
		
		// dyn. Menues definieren
		List<MCommand>commands = application.getCommands();
		for(MCommand command : commands)
		{
			if(StringUtils.equals(command.getElementId(),NEW_DRAW_COMMAND_ID))
				newMenu.addHandledDynamicItem(NEW_DRAW_MENUE_ID,NEW_DRAW_LABEL,command,DRAW_MENUE_POSITION);				

			if(StringUtils.equals(command.getElementId(),NEW_CALC_COMMAND_ID))
				newMenu.addHandledDynamicItem(NEW_CALC_MENUE_ID,NEW_CALC_LABEL,command,DRAW_MENUE_POSITION+1);
			
			if(StringUtils.equals(command.getElementId(),NEW_TEXT_COMMAND_ID))
				newMenu.addHandledDynamicItem(NEW_TEXT_MENUE_ID,NEW_TEXT_LABEL,command,DRAW_MENUE_POSITION+2);			
		}
	}

}
