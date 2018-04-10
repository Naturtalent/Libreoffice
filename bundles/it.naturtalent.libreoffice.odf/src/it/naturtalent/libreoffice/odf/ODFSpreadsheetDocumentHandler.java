package it.naturtalent.libreoffice.odf;

import java.io.File;

import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.TextDocument;

public class ODFSpreadsheetDocumentHandler extends ODFOfficeDocumentHandler
{
	private SpreadsheetDocument odfSpreadsheetDocument;
	
	@Override
	public void openOfficeDocument(File fileDoc)
	{
		try
		{
			this.fileDoc = fileDoc;
			odfSpreadsheetDocument = SpreadsheetDocument.loadDocument(fileDoc);
			
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	@Override
	public void saveOfficeDocument()
	{
		if (odfSpreadsheetDocument != null)
		{
			try
			{
				odfSpreadsheetDocument.save(fileDoc);
			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	@Override
	public void closeOfficeDocument()
	{
		if (odfSpreadsheetDocument != null)
		{
			odfSpreadsheetDocument.close();
			odfSpreadsheetDocument = null;
		}
	}
	
	@Override
	public Object getDocument()
	{		
		return odfSpreadsheetDocument;
	}

}
