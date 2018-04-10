package it.naturtalent.libreoffice;

import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.internal.workbench.E4Workbench;
import org.eclipse.e4.ui.model.application.MApplication;
import org.eclipse.e4.ui.workbench.IWorkbench;

import com.sun.star.lang.EventObject;
import com.sun.star.view.XSelectionChangeListener;

public class ShapeSelectionListener implements XSelectionChangeListener
{

	/*
	 * 
	 * !!! kein Breakpoint in Debug - blockiert das Gesamtsystem
	 *  
	 */
	
	
	@Override
	public void disposing(EventObject arg0)
	{
		// TODO Auto-generated method stub
		
	}

	@Override
	public void selectionChanged(EventObject arg0)
	{		
		// EventBroker informiert, dass Ladevorgang abgeschlossen ist
		MApplication currentApplication = E4Workbench.getServiceContext().get(IWorkbench.class).getApplication();
		IEventBroker eventBroker = currentApplication.getContext().get(IEventBroker.class);
		eventBroker.post(DrawDocumentEvent.DRAWDOCUMENT_EVENT_SHAPE_SELECTED, (Object) arg0);
	}

}
