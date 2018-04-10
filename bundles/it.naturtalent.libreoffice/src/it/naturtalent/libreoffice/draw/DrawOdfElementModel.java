package it.naturtalent.libreoffice.draw;

import it.naturtalent.libreoffice.odf.ODFDrawDocumentHandler;
import it.naturtalent.libreoffice.odf.shapes.IDrawOdfElementModel;

import java.io.File;
import java.util.List;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.odftoolkit.odfdom.dom.element.draw.DrawPageElement;
import org.odftoolkit.odfdom.pkg.OdfElement;

public class DrawOdfElementModel implements IDrawOdfElementModel
{
	private ODFDrawDocumentHandler odfDocumentHandler  = new ODFDrawDocumentHandler();

	public List<OdfElement> getOdfElements(String layerName)
	{
		return getOdfElements(null, layerName);
	}

	public List<OdfElement> getOdfElements(String pageName, String layerName)
	{
		return odfDocumentHandler.getPageLayerElements(pageName, layerName);
	}
	
	public void openDrawDocument(String sourceDrawDocumentPath)
	{
		closeDrawDocument();
		File file = new File(sourceDrawDocumentPath);
		if(file.exists())				
			odfDocumentHandler.openOfficeDocument(file);	
	}
	
	public void closeDrawDocument()
	{
		if(odfDocumentHandler != null)
			odfDocumentHandler.closeOfficeDocument();			
	}

	@Override
	public ODFDrawDocumentHandler getOdfDrawDocumentHandler()
	{	
		return odfDocumentHandler;
	}
	
	
	
}
