package it.naturtalent.libreoffice.odf.shapes;

import java.util.List;

import org.odftoolkit.odfdom.pkg.OdfElement;

import it.naturtalent.libreoffice.odf.ODFDrawDocumentHandler;

public interface IDrawOdfElementModel
{
	public List<OdfElement>getOdfElements(String layer);
	public List<OdfElement>getOdfElements(String drawPage, String layer);
	public void openDrawDocument(String sourceDrawDocumentPath);
	public void closeDrawDocument();
	public ODFDrawDocumentHandler getOdfDrawDocumentHandler();
}
