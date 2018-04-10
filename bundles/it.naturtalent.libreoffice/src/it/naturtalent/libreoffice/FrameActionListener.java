package it.naturtalent.libreoffice;

import org.eclipse.e4.core.services.events.IEventBroker;
import org.eclipse.e4.ui.internal.workbench.E4Workbench;
import org.eclipse.e4.ui.model.application.MApplication;
import org.eclipse.e4.ui.workbench.IWorkbench;

import com.sun.star.frame.FrameAction;
import com.sun.star.frame.FrameActionEvent;
import com.sun.star.frame.XFrameActionListener;
import com.sun.star.lang.EventObject;

public class FrameActionListener implements XFrameActionListener
{
	@Override
	public void disposing(EventObject arg0)
	{
	}

	@Override
	public void frameAction(FrameActionEvent aEvent)
	{
	    if(aEvent.Action.getValue() == FrameAction.FRAME_ACTIVATED_value)
	    {	    
	    	MApplication currentApplication = E4Workbench.getServiceContext().get(IWorkbench.class).getApplication();
	    	IEventBroker eventBroker = currentApplication.getContext().get(IEventBroker.class);	    		    	
	    	eventBroker.post(DrawDocumentEvent.DRAWDOCUMENT_EVENT_DOCUMENT_ACTIVATE, aEvent.Frame);		
	    }
	}

}
