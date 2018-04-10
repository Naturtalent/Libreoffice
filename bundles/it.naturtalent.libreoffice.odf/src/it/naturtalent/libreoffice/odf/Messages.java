package it.naturtalent.libreoffice.odf;

import org.eclipse.osgi.util.NLS;

public class Messages extends NLS
{
	private static final String BUNDLE_NAME = "it.naturtalent.e4.office.odf.messages"; //$NON-NLS-1$

	public static String ODFOfficeDocumentHandler_3;

	public static String ODFOfficeDocumentHandler_CANNOT_START_ERROR;

	public static String ODFOfficeDocumentHandler_NO_OFFICE_ERROR;

	public static String ODFOfficeDocumentHandler_OFFICEPATH;
	
	public static String ODFOfficeDocumentHandler_OPEN_ERROR;
	
	public static String ODFOfficeDocumentHandler_OFFICE_APPLICATION_PATH_ERROR;
	
	static
	{
		// initialize resource bundle
		NLS.initializeMessages(BUNDLE_NAME, Messages.class);
	}

	private Messages()
	{
	}
}
