package it.naturtalent.libreoffice;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import it.naturtalent.application.services.IOpenWithEditorAdapter;
import it.naturtalent.libreoffice.calc.CalcDocument;
import it.naturtalent.libreoffice.draw.DrawDocument;
import it.naturtalent.libreoffice.text.TextDocument;
import it.naturtalent.libreoffice.utils.Lo;

/**
 * Nutzt das OpenWith - Interface zum Öffnen von Lo Dokumenten. 
 * Dieser Adapter bindet keinen spezifischen Editor via dyn. Menue ein, sondern oeffnet alle Lo-Dokumente mit Libreoffice.
 * 
 * @author dieter
 *
 */
public class OpenLoDocumentAdapter implements IOpenWithEditorAdapter
{
	// Array der betroffenen Dateien
	//private static String [] loExtensins = {"odt","odg","ods"};

	
	@Override
	public String getCommandID()
	{		
		return null; // kein spezifischer Handler
	}

	@Override
	public String getMenuID()
	{		
		return null;		
	}

	@Override
	public String getMenuLabel()
	{		
		return "LibreOffice";		
	}

	
	@Override
	public String getContribURI()
	{
		return null;
		//return "bundleclass://it.naturtalent.e4.project.ui/it.naturtalent.e4.project.ui.handlers.TESThandler";		
	}

	@Override
	public boolean getType()
	{		
		return true;
	}

	@Override
	public int getIndex()
	{		
		return (-1);
	}

	// ausfuehrbar, wenn der Dateiname in 'filePath' eine Lo-Datei ist 
	@Override
	public boolean isExecutable(String filePath)
	{
		String ext = FilenameUtils.getExtension(filePath);	
		return (Lo.ext2DocType(ext) != null);
	}

	// mit Libreoffice oeffnen
	@Override
	public void execute(String filePath)
	{
		if(isExecutable(filePath))
		{
			OpenLoDocument.loadLoDocument(filePath);
		}
	}

}
