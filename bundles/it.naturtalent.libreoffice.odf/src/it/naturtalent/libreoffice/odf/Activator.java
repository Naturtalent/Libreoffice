package it.naturtalent.libreoffice.odf;



import java.io.File;
import java.net.URL;

import org.eclipse.core.runtime.FileLocator;
import org.eclipse.core.runtime.Path;
import org.osgi.framework.BundleActivator;
import org.osgi.framework.BundleContext;

public class Activator implements BundleActivator
{
	
	public static final String PLUGIN_TEMPLATE_DIR = File.separator + "templates"; //$NON-NLS-1$
	

	private static BundleContext context;

	static BundleContext getContext()
	{
		return context;
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see
	 * org.osgi.framework.BundleActivator#start(org.osgi.framework.BundleContext
	 * )
	 */
	public void start(BundleContext bundleContext) throws Exception
	{
		Activator.context = bundleContext;		
		ODFPreferences.initialize();
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see
	 * org.osgi.framework.BundleActivator#stop(org.osgi.framework.BundleContext)
	 */
	public void stop(BundleContext bundleContext) throws Exception
	{
		Activator.context = null;
	}
	
	public static URL getPluginPath(String path)
	{
		try
		{
			URL url = FileLocator.find(Activator.context.getBundle(), new Path(path), null);
			url = FileLocator.resolve(url);
			return url;
		
		} catch (Exception e)
		{
		}
		return null;
	}


}
