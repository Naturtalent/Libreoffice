package it.naturtalent.libreoffice.test;

import com.sun.star.awt.XItemListener;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.EventObject;
import com.sun.star.uno.UnoRuntime;

public class RoadmapItemStateChangeListener implements XItemListener
{
	protected com.sun.star.lang.XMultiServiceFactory m_xMSFDialogModel;

	public RoadmapItemStateChangeListener(
			com.sun.star.lang.XMultiServiceFactory xMSFDialogModel)
	{
		m_xMSFDialogModel = xMSFDialogModel;
	}

	public void itemStateChanged(com.sun.star.awt.ItemEvent itemEvent)
	{
		try
		{
			// get the new ID of the roadmap that is supposed to refer to the
			// new step of the dialogmodel
			int nNewID = itemEvent.ItemId;
			XPropertySet xDialogModelPropertySet = (XPropertySet) UnoRuntime
					.queryInterface(XPropertySet.class, m_xMSFDialogModel);
			int nOldStep = ((Integer) xDialogModelPropertySet
					.getPropertyValue("Step")).intValue();
			// in the following line "ID" and "Step" are mixed together.
			// In fact in this case they denot the same
			if (nNewID != nOldStep)
			{
				xDialogModelPropertySet.setPropertyValue("Step", new Integer(
						nNewID));
			}
		} catch (com.sun.star.uno.Exception exception)
		{
			exception.printStackTrace(System.out);
		}
	}

	@Override
	public void disposing(EventObject arg0)
	{
		// TODO Auto-generated method stub

	}
}
