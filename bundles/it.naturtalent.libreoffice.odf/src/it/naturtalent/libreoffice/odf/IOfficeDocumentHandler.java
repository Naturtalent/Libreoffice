package it.naturtalent.libreoffice.odf;

import java.io.File;

/**
 * @author apel.dieter
 *
 */
public interface IOfficeDocumentHandler
{
	public static final String ODF_OFFICETEXTDOCUMENT_EXTENSION = "odt";
	public static final String MS_OFFICETEXTDOCUMENT_EXTENSION = "docx";
	
	
	public void openOfficeDocument(File fileDoc);
	public void saveOfficeDocument();
	public void closeOfficeDocument();
	public void showOfficeDocument();
	public Object getDocument();
	public String getProperty();
	public void setProperty(String property);
}
