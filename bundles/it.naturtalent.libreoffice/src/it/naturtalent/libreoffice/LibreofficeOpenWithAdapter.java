package it.naturtalent.libreoffice;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;

import it.naturtalent.application.services.IOpenWithEditorAdapter;
import it.naturtalent.libreoffice.calc.CalcDocument;
import it.naturtalent.libreoffice.draw.DrawDocument;
import it.naturtalent.libreoffice.text.TextDocument;

/**
 * Dieser Adapter bindet kein dyn. Menue ein, sondern oeffnet alle Files mit (odt,odd,..) mit LibreOffice.
 * 
 * @author dieter
 *
 */
public class LibreofficeOpenWithAdapter implements IOpenWithEditorAdapter
{

	private static String [] extensions = {"odt","odg","ods"};
	private static enum FileExtension
	{
		TextFile,
		DrawFile,
		CalcFile;
				
		public static String getExtension(FileExtension fileExtension)
		{
			String ext ="";
			switch (fileExtension)
				{
					case TextFile:
						return extensions[0];
						
					case DrawFile:
						return extensions[1];

					case CalcFile:
						return extensions[2];
				}
			return ext;
		}
	}
	
	
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

	// ausfuehrbar, wenn Dateiname mit 'odt' Extension 
	@Override
	public boolean isExecutable(String filePath)
	{
		String ext = FilenameUtils.getExtension(filePath);		
		return (StringUtils.equalsAny(ext, extensions[0],extensions[1],extensions[2]));		
	}

	// ODT-Textdatei mit Libreoffice oeffnen
	@Override
	public void execute(String filePath)
	{
		String ext = FilenameUtils.getExtension(filePath);	
		
		if(StringUtils.equals(ext, FileExtension.getExtension(FileExtension.TextFile)))
		{
			TextDocument txtDocument = new TextDocument();
			txtDocument.loadPage(filePath);	
			return;
		}
		else
		{
			if(StringUtils.equals(ext, FileExtension.getExtension(FileExtension.DrawFile)))
			{
				DrawDocument drawDocument = new DrawDocument();
				drawDocument.loadPage(filePath);
				return;
			}
			else
			{
				if(StringUtils.equals(ext, FileExtension.getExtension(FileExtension.CalcFile)))
				{
					CalcDocument calcDocument = new CalcDocument();
					calcDocument.loadPage(filePath);
					return;
				}				
			}
		}
	}

}
