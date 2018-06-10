package it.naturtalent.libreoffice.odf.text;

import java.io.File;

import org.odftoolkit.simple.TextDocument;

public class ODFTextHandler
{	
	private TextDocument odfTextDocument;
	
	/**
	 * 
	 * @param fileDoc
	 */
	public void openODFTextDocument(File fileDoc)
	{
		try
		{
			odfTextDocument = TextDocument.loadDocument(fileDoc);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
