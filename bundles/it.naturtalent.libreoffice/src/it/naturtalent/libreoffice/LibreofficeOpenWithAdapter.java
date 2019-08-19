package it.naturtalent.libreoffice;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;

import it.naturtalent.application.services.IOpenWithEditorAdapter;
import it.naturtalent.libreoffice.calc.CalcDocument;
import it.naturtalent.libreoffice.draw.DrawDocument;
import it.naturtalent.libreoffice.text.TextDocument;
import it.naturtalent.libreoffice.utils.Lo;

/**
 * Mit diesem Adapter koennen Lo-Dateien geoffnet werden. 
 * @see it.naturtalent.e4.project.ui.actions,SystenOpenEditorAction
 * 
 * @author dieter
 *
 */
public class LibreofficeOpenWithAdapter implements IOpenWithEditorAdapter
{
	@Override
	public String getCommandID()
	{		
		return null;
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

	// ausfuehrbar, wenn Extension einer Lo-Datei entspricht 
	@Override
	public boolean isExecutable(String filePath)
	{
		String ext = FilenameUtils.getExtension(filePath);	
		
		String docType = Lo.ext2DocType(ext);
		if(StringUtils.equals(docType, Lo.WRITER_STR))
			return StringUtils.equals(ext, "odt") ? true : false;
		
		return true;		
	}

	// ODT-Textdatei mit Libreoffice oeffnen
	@Override
	public void execute(String filePath)
	{		
		OpenLoDocument.loadLoDocument(filePath);
	}

}
