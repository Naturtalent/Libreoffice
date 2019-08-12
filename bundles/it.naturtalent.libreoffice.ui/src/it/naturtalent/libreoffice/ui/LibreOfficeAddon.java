 
package it.naturtalent.libreoffice.ui;

import javax.inject.Inject;

import org.apache.commons.lang3.SystemUtils;
import org.eclipse.core.runtime.preferences.DefaultScope;
import org.eclipse.core.runtime.preferences.IEclipsePreferences;
import org.eclipse.e4.core.di.annotations.Optional;
import org.eclipse.e4.core.di.extensions.EventTopic;
import org.eclipse.e4.ui.workbench.UIEvents;
import org.osgi.service.event.Event;

import it.naturtalent.e4.preferences.IPreferenceRegistry;

/**
 * AddOn - Initialisierung, definiert in 'fragment.e4xmi' 
 * 
 * @author dieter
 *
 */
public class LibreOfficeAddon 
{
	
	private @Inject @Optional IPreferenceRegistry preferenceRegistry;

	@Inject
	@Optional
	public void applicationStarted(@EventTopic(UIEvents.UILifeCycle.APP_STARTUP_COMPLETE) Event event)
	{
		/* 
		 * Defaultpreferences LibreOffice
		 * 
		 */
		IEclipsePreferences defaultPreferences = DefaultScope.INSTANCE
				.getNode(OfficeConstants.ROOT_OFFICE_PREFERENCES_NODE);
	
		if (SystemUtils.IS_OS_LINUX)
		{
			// Linux speichert Libreoffice in einem definierten Speicher (OfficeConstants.LINUX_UNO_PATH)
			defaultPreferences.put(OfficeConstants.OFFICE_APPLICATION_PREF,OfficeConstants.LINUX_UNO_PATH);
		}
		else
		{
			// in Windows gibt es kein definiertes Verzeicnis für Libreoffice, somit auch kein Defaultverzeicnis 
		}
		
		// der LibreOffice - Prefaerenzadapter in der Registry speichern 
		if(preferenceRegistry != null)	
			preferenceRegistry.getPreferenceAdapters().add(new LibreofficeApplicationPreferenceAdapter());
	}

}
