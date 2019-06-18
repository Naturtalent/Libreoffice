package it.naturtalent.libreoffice;

import java.io.File;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.SystemUtils;
import org.eclipse.core.runtime.preferences.DefaultScope;
import org.eclipse.core.runtime.preferences.IEclipsePreferences;
import org.eclipse.core.runtime.preferences.InstanceScope;

import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XBridge;
import com.sun.star.bridge.XBridgeFactory;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.connection.XConnection;
import com.sun.star.connection.XConnector;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.XComponentContext;

import it.naturtalent.libreoffice.utils.Lo;

/**
 * 
 * @author dieter
 *
 */
public class SocketContext
{
	// connect to locally running Office via port 8100
	private static final int SOCKET_PORT = 8100;  
	
	// bridge ermoeglicht die prozessuebergreifende Kommunikation mit Libreoffice  
	// bei einer Socket-Connection
	private static XBridge bridge = null;
	
	/**
	 * Libreoffice ueber eine 
	 * @return
	 */
	public static XComponentContext socketContext()
	{
		try
		{
			if (bridge != null)
				return getRemoteContext();
			
			IEclipsePreferences preferences = InstanceScope.INSTANCE
					.getNode(OfficeConstants.ROOT_OFFICE_PREFERENCES_NODE);
			String officApplicationPath = preferences.get(OfficeConstants.OFFICE_APPLICATION_PREF, null);
			
			if(StringUtils.isEmpty(officApplicationPath))
			{
				preferences = DefaultScope.INSTANCE
						 .getNode(OfficeConstants.ROOT_OFFICE_PREFERENCES_NODE);
				officApplicationPath = preferences.get(OfficeConstants.OFFICE_APPLICATION_PREF, null);
			}
			
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
			
			String[] cmdArray = new String[3];
			cmdArray[0] = fOffice.getPath();			
			cmdArray[1] = "-headless";
			cmdArray[2] = "-accept=socket,host=localhost,port=" + SOCKET_PORT + ";urp;";
			Process p = Runtime.getRuntime().exec(cmdArray);
			if (p != null)
				System.out.println("Office process created");
			Lo.delay(5000);
			// Wait 5 seconds, until office is in listening mode

			// Create a local Component Context
			XComponentContext localContext = Bootstrap.createInitialComponentContext(null);

			// Get the local service manager
			XMultiComponentFactory localFactory = localContext.getServiceManager();

			XConnector connector = Lo.qi(XConnector.class,
					localFactory.createInstanceWithContext(
							"com.sun.star.connection.Connector", localContext));

			XConnection connection = connector
					.connect("socket,host=localhost,port=" + SOCKET_PORT);

			// create a bridge to Office via the socket
			XBridgeFactory bridgeFactory = Lo.qi(XBridgeFactory.class,
					localFactory.createInstanceWithContext(
							"com.sun.star.bridge.BridgeFactory", localContext));

			bridge = bridgeFactory.createBridge("socketBridgeAD", "urp",connection, null);
			XComponent bridgeComponent = Lo.qi(XComponent.class, bridge);
			bridgeComponent.addEventListener(new XEventListener()
			{				
				@Override
				public void disposing(EventObject arg0)
				{
					// Socketconnection ist definitiv unterbrochen (Lo crashed or terminated)
					bridge = null;
					System.out.println("Bridge disposed");					
				}
			});
			
			return getRemoteContext();
			
		} catch (java.lang.Exception e)
		{
			System.out.println("Unable to socket connect to Office");
		}

		return null;
	} 
	
	/*
	 * Ueber eine bestehende bridge wird der Context zurueckgegeben.
	 * 
	 * Es kann nicht ueberprueft werden, die bridge real noch existiert 
	 */
	private static XComponentContext getRemoteContext() throws Exception
	{
		// get the remote service manager
		XMultiComponentFactory serviceManager = Lo.qi(
				XMultiComponentFactory.class,
				bridge.getInstance("StarOffice.ServiceManager"));

		// retrieve Office's remote component context as a property
		XPropertySet props = Lo.qi(XPropertySet.class, serviceManager);
		// initObject);
		Object defaultContext = props.getPropertyValue("DefaultContext");

		// get the remote interface XComponentContext
		return Lo.qi(XComponentContext.class, defaultContext);
	}
	
	public static void terminateConnection()
	{
		bridge = null;
	}

}
