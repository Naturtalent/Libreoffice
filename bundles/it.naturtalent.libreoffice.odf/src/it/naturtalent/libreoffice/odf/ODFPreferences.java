package it.naturtalent.libreoffice.odf;

import java.io.File;
import java.net.URL;

import it.naturtalent.e4.project.ui.NtPreferences;

import org.apache.commons.lang3.StringUtils;
import org.eclipse.core.runtime.preferences.DefaultScope;
import org.eclipse.core.runtime.preferences.IEclipsePreferences;

public class ODFPreferences
{

	public static void initialize()
	{
		StringBuilder builder = new StringBuilder();
		
		URL url = Activator.getPluginPath(Activator.PLUGIN_TEMPLATE_DIR+File.separator+"calc.ods");		
		builder.append("ODFCalc,"+url.toExternalForm()+",");
		url = Activator.getPluginPath(Activator.PLUGIN_TEMPLATE_DIR+File.separator+"draw.odg");
		builder.append("ODFDraw,"+url.toExternalForm());
		
		IEclipsePreferences defaultNode = DefaultScope.INSTANCE
				.getNode(NtPreferences.ROOT_PREFERENCES_NODE);
		String defaultValue = defaultNode.get(NtPreferences.FILE_TEMPLATE_PREFERENCE, null);
		String value = builder.toString();
		defaultValue = StringUtils.isNotEmpty(defaultValue) ? defaultValue
				+ ","+value : value; 
		defaultNode.put(NtPreferences.FILE_TEMPLATE_PREFERENCE, defaultValue);
	}
}
