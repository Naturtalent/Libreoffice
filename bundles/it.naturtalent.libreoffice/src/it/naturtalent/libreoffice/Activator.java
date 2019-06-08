package it.naturtalent.libreoffice;

import org.apache.commons.lang3.SystemUtils;
import org.eclipse.ui.plugin.AbstractUIPlugin;
import org.jnativehook.GlobalScreen;
import org.osgi.framework.BundleContext;

import com.sun.star.uno.XComponentContext;

import it.naturtalent.libreoffice.utils.Lo;

/**
 * The activator class controls the plug-in life cycle
 */
public class Activator extends AbstractUIPlugin
{

	// The plug-in ID
	public static final String PLUGIN_ID = "it.naturtalent.libreoffice"; //$NON-NLS-1$

	// The shared instance
	private static Activator plugin;

	/**
	 * The constructor
	 */
	public Activator()
	{
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see org.eclipse.ui.plugin.AbstractUIPlugin#start(org.osgi.framework.
	 * BundleContext)
	 */
	public void start(BundleContext context) throws Exception
	{
		super.start(context);
		plugin = this;
		//startLibreoffice();
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see org.eclipse.ui.plugin.AbstractUIPlugin#stop(org.osgi.framework.
	 * BundleContext)
	 */
	public void stop(BundleContext context) throws Exception
	{		
		plugin = null;		
	}

	/**
	 * Returns the shared instance
	 *
	 * @return the shared instance
	 */
	public static Activator getDefault()
	{
		return plugin;
	}

}
