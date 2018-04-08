package it.naturtalent.libreoffice.ui;

public class OfficeConstants
{
	// Plugin ID (MANIFEST.MF)
	public static final String PLUGIN_ID = "it.naturtalent.e4.office.ui"; //$NON-NLS-1$
	
	// Office Preferencenode
	public static final String ROOT_OFFICE_PREFERENCES_NODE = "it.naturtalent.e4.office"; //$NON-NLS-1$
	
	// Praeferenzkey zum Pfad der Officeapplication
	public static final String OFFICE_APPLICATION_PREF = "officeapplication_pref"; //$NON-NLS-1$
	
	// Praeferenzkey zum Pfad der JPIPE-Library
	public static final String OFFICE_JPIPE_PREF = "officejpipe_pref"; //$NON-NLS-1$

	// Praeferenzkey zum Pfad der UNO-Komponenten
	public static final String OFFICE_UNO_PREF = "officeuno_pref"; //$NON-NLS-1$

	
	// plugin.properties
	public static final String PROPERTY_OFFICENEWLETTER = "Office.NewLetterLabel"; //$NON-NLS-1$

	// externe Officeanwendung (LibreOffice) @see OfficeProcessor()
	public static final String LINUX_UNO_PATH = "/usr/lib/libreoffice/program/";
	public static final String WINDOWS_UNO_PATH = "\\Office\\LibreOfficeProtable4\\App\\libreoffice\\program"; //$NON-NLS-1$	

	// Settings
	//public static final String NTOFFICE_SETTINGTEMPLATE_KEY = "ntofficetemplatesetting"; //$NON-NLS-1$
	public static final String NTOFFICE_SETTINGLETTERFILENAME_KEY = "ntofficeletterdocumentsetting"; //$NON-NLS-1$
}
