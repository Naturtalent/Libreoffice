package it.naturtalent.libreoffice.calc;

import java.io.File;

import org.apache.commons.lang.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Row;
import org.odftoolkit.simple.table.Table;

import it.naturtalent.libreoffice.odf.ODFSpreadsheetDocumentHandler;


/**
 * ODF Tabellenadapter
 * 
 * @author dieter
 *
 */
public class ODFSpreadsheetAdapter extends AbstractSpreadSheetAdapter
{
	
	private ODFSpreadsheetDocumentHandler documentHandler = new ODFSpreadsheetDocumentHandler();
	private SpreadsheetDocument document;

	private Table table; 		
	 		
		
	@Override
	public void openSpreadsheetDocument(File docFile)
	{
		documentHandler.openOfficeDocument(docFile);
		document = (SpreadsheetDocument) documentHandler.getDocument();
	}

	@Override
	public boolean openSheet(String tableName)
	{
		table = null;
		if(document != null)
		{
			if(StringUtils.isNotEmpty(tableName))
				table = document.getSheetByName(tableName);
			else
				table = document.getSheetByIndex(0);
		}
		
		return (table != null);
	}
	
	@Override
	public void appendSheet(String tableName)
	{
		table = null;
		if(document != null)
		{
			if(StringUtils.isNotEmpty(tableName))
				table = document.appendSheet(tableName);
		}
	}

	@Override
	public void cloneSheet(int idx, String name)
	{
		Table table = document.getSheetByIndex(idx);
		table = document.insertSheet(table,idx);
		table.setTableName(name);
	}

	
	@Override
	public void renameSheet(String sheetName)
	{
		if(table != null)
			table.setTableName(sheetName);
	}

	@Override
	public String [] readRow(int rowIdx, int nCols)
	{
		String [] values = null;
		if (table != null)
		{
			Row row = table.getRowByIndex(rowIdx);			
			for(int i = 0;i < nCols;i++)
			{
				Cell cell = row.getCellByIndex(i);
				values = (String[]) ArrayUtils.add(values, cell.getStringValue());
			}			
		}
		return values;
	}
	
	@Override
	public void writeRow(String [] values, int rowIdx, int colIdx)
	{
		if (table != null)
		{
			Row row = table.getRowByIndex(rowIdx);
			for(int i = 0;i < values.length;i++)
			{
				Cell cell = row.getCellByIndex(i+colIdx);
				cell.setStringValue(values[i]);
			}
		}		
	}
	
	@Override
	public void closeSpreadsheetDocument()
	{
		documentHandler.saveOfficeDocument();		
		documentHandler.closeOfficeDocument();
	}

	@Override
	public void showSpreadsheetDocument()
	{
		documentHandler.showOfficeDocument();		
	}

}
