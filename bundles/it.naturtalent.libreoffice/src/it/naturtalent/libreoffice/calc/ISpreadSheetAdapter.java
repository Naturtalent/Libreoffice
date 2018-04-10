package it.naturtalent.libreoffice.calc;

import java.io.File;

/**
 * Der Adapter ermoeglicht die Bearbeitung unterschiedlicher Tabellen (ODF und MS) mit einheitlichen Funktionen.
 * 
 * @author dieter
 *
 */
public interface ISpreadSheetAdapter
{
	// Namen, unter denen die Defaultadapter im Registry 'OfficeService' gesoeichert sind 
	public static final String DEFAULT_SPREADSHEET_ADAPTER = "defaultspreadsheetadapter";
	public static final String DEFAULT_MSSPREADSHEET_ADAPTER = "defaultmsspreadsheetadapter";
	
	public void openSpreadsheetDocument(File docFile);
	
	public boolean openSheet(String tableName);
	public void appendSheet(String tableName);
	public void cloneSheet(int srcIdx, String name);
	public void renameSheet(String sheetName);
	public String [] readRow(int rowIdx, int nCols);
	public void writeRow(String [] values, int rowIdx, int colIdx);
	
	public void closeSpreadsheetDocument();
	public void showSpreadsheetDocument();

		
}
