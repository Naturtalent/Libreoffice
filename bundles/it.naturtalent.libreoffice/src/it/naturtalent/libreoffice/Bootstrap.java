package it.naturtalent.libreoffice;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.util.Hashtable;
import java.util.Random;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.SystemUtils;
import org.apache.commons.logging.LogFactory;
import org.eclipse.core.runtime.preferences.DefaultScope;
import org.eclipse.core.runtime.preferences.IEclipsePreferences;
import org.eclipse.core.runtime.preferences.InstanceScope;
import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.internal.workbench.E4Workbench;
import org.eclipse.e4.ui.model.application.MApplication;
import org.eclipse.e4.ui.workbench.IWorkbench;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.swt.widgets.Display;

import com.sun.star.bridge.UnoUrlResolver;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.comp.helper.ComponentContextEntry;
import com.sun.star.comp.loader.JavaLoader;
import com.sun.star.container.XSet;
import com.sun.star.lang.XInitialization;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.loader.XImplementationLoader;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;



public class Bootstrap
{

	//static String OFFICE = "C:\\Users\\A682055\\Daten\\Naturtalent4\\programme\\Office\\LibreOfficeProtable4\\App\\libreoffice\\program";
	//static String LINUX_OFFICE = "/usr/lib/libreoffice/program/";
	
	public static final XComponentContext bootstrap() throws BootstrapException
	{
		XComponentContext xLocalContext = null;
		XComponentContext xContext = null;
		XMultiComponentFactory xLocalServiceManager = null;
		
		try
		{			
			IEclipsePreferences preferences = InstanceScope.INSTANCE
					.getNode(OfficeConstants.ROOT_OFFICE_PREFERENCES_NODE);
			String officApplicationPath = preferences.get(OfficeConstants.OFFICE_APPLICATION_PREF, null);
			
			if(StringUtils.isEmpty(officApplicationPath))
			{
				preferences = DefaultScope.INSTANCE
						.getNode(OfficeConstants.ROOT_OFFICE_PREFERENCES_NODE);
				officApplicationPath = preferences.get(OfficeConstants.OFFICE_APPLICATION_PREF, null);
			}
			
			// create default local component context
			xLocalContext = createInitialComponentContext(null);
			if (xLocalContext == null)
				throw new BootstrapException("no local component context!");
			
			String sOffice;
			if(SystemUtils.IS_OS_WINDOWS)
				sOffice = officApplicationPath +File.separator+"soffice.exe";
			else
			{
				//officApplicationPath = (StringUtils.isEmpty(officApplicationPath)) ? LINUX_OFFICE : officApplicationPath; 					
				sOffice = officApplicationPath +File.separator+"soffice";
			}
			
			File fOffice = new File(sOffice);
			if(!fOffice.exists())
				throw new BootstrapException("no office executable found!");
			
			// create random pipe name
			String sPipeName = "uno" + Long.toString((new Random()).nextLong() & 0x7fffffffffffffffL);
			
			// create call with arguments
			String[] cmdArray = new String[7];
			cmdArray[0] = fOffice.getPath();
			cmdArray[1] = "--nologo";
			cmdArray[2] = "--nodefault";
			cmdArray[3] = "--norestore";
			cmdArray[4] = "--nocrashreport";
			cmdArray[5] = "--nolockcheck";
			cmdArray[6] = "--accept=pipe,name=" + sPipeName + ";urp;";
			// start office process
			Process p = Runtime.getRuntime().exec(cmdArray);
			pipe(p.getInputStream(), System.out, "CO> ");
			pipe(p.getErrorStream(), System.err, "CE> ");
			// initial service manager
			xLocalServiceManager = xLocalContext.getServiceManager();
			if (xLocalServiceManager == null)
				throw new BootstrapException("no initial service manager!");
			
			// create a URL resolver
			XUnoUrlResolver xUrlResolver = UnoUrlResolver.create(xLocalContext);
			// connection string
			String sConnect = "uno:pipe,name=" + sPipeName
					+ ";urp;StarOffice.ComponentContext";
			// wait until office is started
			for (int i = 0;; ++i)
			{
				try
				{
					// try to connect to office
					Object context = xUrlResolver.resolve(sConnect);
					xContext = UnoRuntime.queryInterface(
							XComponentContext.class, context);
					if (xContext == null)
						throw new BootstrapException("no component context!");					
					break;
				} catch (com.sun.star.connection.NoConnectException ex)
				{
					// Wait 500 ms, then try to connect again, but do not wait
					// longer than 5 min (= 600 * 500 ms) total:
					if (i == 600)
					{
						throw new BootstrapException(ex.toString());
					}
					Thread.sleep(500);
				}
			}
		} catch (final Exception e)
		{
			e.printStackTrace();
			/*
			Display.getDefault().syncExec(new Runnable()
			{
				public void run()
				{
					LogFactory.getLog(this.getClass()).error(e);
					
					// Watchdog (@see OpenDesignAction) abschalten
					MApplication currentApplication = E4Workbench.getServiceContext().get(IWorkbench.class).getApplication();
					IEventBroker eventBroker = currentApplication.getContext().get(IEventBroker.class);
					eventBroker.post(DrawDocumentEvent.DRAWDOCUMENT_EVENT_DOCUMENT_OPEN_CANCEL, null);
					
					MessageDialog
							.openError(
									Display.getDefault()
											.getActiveShell(),
									"LibreOffice",
									"Fehler beim Zugriff auf LibreOffice");
				}
			});
			*/				
		}
		
		
		return xContext;
	}
	
	 
	 /**
	* backwards compatibility stub.
	*/
	static public XComponentContext createInitialComponentContext(
			Hashtable<String, Object> context_entries) throws Exception
	{
		return createInitialComponentContext((java.util.Map<String, Object>) context_entries);
	}
	

	/**
	 * Bootstraps an initial component context with service manager and basic
	 * jurt components inserted.
	 * 
	 * @param context_entries
	 *            the hash table contains mappings of entry names (type string)
	 *            to context entries (type class ComponentContextEntry).
	 * @return a new context.
	 */
	
	static public XComponentContext createInitialComponentContext(
			java.util.Map<String, Object> context_entries) throws Exception
	{
		ServiceManager xSMgr = new ServiceManager();
		XImplementationLoader xImpLoader = UnoRuntime.queryInterface(
				XImplementationLoader.class, new JavaLoader());
		XInitialization xInit = UnoRuntime.queryInterface(
				XInitialization.class, xImpLoader);
		Object[] args = new Object[]{ xSMgr };
		xInit.initialize(args);
		
		// initial component context
		if (context_entries == null)
			context_entries = new Hashtable<String, Object>(1);
		
		// add smgr
		context_entries.put("/singletons/com.sun.star.lang.theServiceManager",
				new ComponentContextEntry(null, xSMgr));
		// ... xxx todo: add standard entries
		
		XComponentContext xContext = new ComponentContext(context_entries, null);
		xSMgr.setDefaultContext(xContext);
		
		XSet xSet = UnoRuntime.queryInterface(XSet.class, xSMgr);
		// insert basic jurt factories
		insertBasicFactories(xSet, xImpLoader);
		
		return xContext;
	}

	private static void insertBasicFactories(XSet xSet,
			XImplementationLoader xImpLoader) throws Exception
	{
		// insert the factory of the loader
		xSet.insert(xImpLoader.activate("com.sun.star.comp.loader.JavaLoader",
				null, null, null));
		// insert the factory of the URLResolver
		xSet.insert(xImpLoader.activate(
				"com.sun.star.comp.urlresolver.UrlResolver", null, null, null));
		// insert the bridgefactory
		xSet.insert(xImpLoader.activate(
				"com.sun.star.comp.bridgefactory.BridgeFactory", null, null,
				null));
		// insert the connector
		xSet.insert(xImpLoader.activate(
				"com.sun.star.comp.connections.Connector", null, null, null));
		// insert the acceptor
		xSet.insert(xImpLoader.activate(
				"com.sun.star.comp.connections.Acceptor", null, null, null));
	}
	 
	private static void pipe(final InputStream in, final PrintStream out,
			final String prefix)
	{
		new Thread("Pipe: " + prefix)
		{
			public void run()
			{
				BufferedReader r = new BufferedReader(new InputStreamReader(in));
				try
				{
					for (;;)
					{
						String s = r.readLine();
						if (s == null)
						{
							break;
						}
						out.println(prefix + s);
					}
				} catch (java.io.IOException e)
				{
					e.printStackTrace(System.err);
				}
			}
		}.start();
	}
}
