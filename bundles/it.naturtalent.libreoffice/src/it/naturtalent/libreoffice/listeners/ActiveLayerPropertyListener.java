package it.naturtalent.libreoffice.listeners;

import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.internal.workbench.E4Workbench;
import org.eclipse.e4.ui.model.application.MApplication;
import org.eclipse.e4.ui.workbench.IWorkbench;

import com.sun.star.beans.PropertyChangeEvent;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XLayer;
import com.sun.star.drawing.XLayerManager;
import com.sun.star.drawing.XLayerSupplier;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.Any;
import com.sun.star.uno.RuntimeException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XInterface;

import it.naturtalent.libreoffice.PropertyChangeListenerHelper;


/**
 * Ueberwacht im XPropertySet von XComponent die Eigenschaft mit dem Namen 'ActiveLayer'. Aendert sich der Wert dieser
 * Eigenschaft ist das gleichbdutend mit einer Selektionsaenderung.
 * 
 * @author dieter
 *
 */
public class ActiveLayerPropertyListener extends PropertyChangeListenerHelper
{
	
	// relevante Propertynamen
	private static String [] propertyNames = { "ActiveLayer" };
	
	
	/**
	 * Konstruktion
	 * 
	 * @param xEvtSource (XComponent = DrawCoument)
	 */
	public ActiveLayerPropertyListener(XInterface xEvtSource)
	{
		super(xEvtSource, propertyNames);		
	}

	/**
	 * aktiviert den PageListener des DrawDocumets
	 * 
	 * 
	 * @param xComponent
	 */
	public void activatePageListener()
	{
		XModel xModel = UnoRuntime.queryInterface(XModel.class,GetEvtSource());
		XController xController = xModel.getCurrentController();		
		XPropertySet xPageProperties = UnoRuntime.queryInterface(XPropertySet.class, xController);
		AddAsListenerTo(xPageProperties);
	}

	/**
	 * aktiviert den PageListener des DrawDocumets
	 * 
	 * @param xComponent
	 */
	public void deativatePageListener()
	{
		RemoveAsListener();
	}

	@Override
	public void propertyChange(PropertyChangeEvent aEvt) throws RuntimeException
	{
		//System.out.println("PropertyLayerName");
		
		Object obj = aEvt.NewValue;
		if(obj instanceof XLayer)
		{		
			XPropertySet xLayerProperties = UnoRuntime.queryInterface(XPropertySet.class, obj);
			try
			{
				System.out.println("Property Layer: "+xLayerProperties.getPropertyValue("Name"));
			} catch (UnknownPropertyException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WrappedTargetException e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
}
