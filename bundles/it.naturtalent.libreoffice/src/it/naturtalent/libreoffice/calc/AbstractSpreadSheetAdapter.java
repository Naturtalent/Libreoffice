package it.naturtalent.libreoffice.calc;

import java.io.File;

public abstract class AbstractSpreadSheetAdapter implements ISpreadSheetAdapter
{

	@Override
	public void openSpreadsheetDocument(File docFile)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public boolean openSheet(String tableName)
	{
		// TODO Auto-generated method stub
		return false;
	}

	@Override
	public void cloneSheet(int idx, String name)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void appendSheet(String tableName)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void renameSheet(String sheetName)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public String[] readRow(int rowIdx, int nCols)
	{
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void writeRow(String[] values, int rowIdx, int colIdx)
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void closeSpreadsheetDocument()
	{
		// TODO Auto-generated method stub

	}

	@Override
	public void showSpreadsheetDocument()
	{
		// TODO Auto-generated method stub

	}

}
