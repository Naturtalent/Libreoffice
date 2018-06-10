package it.naturtalent.libreoffice.odf;



import java.io.File;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.odftoolkit.odfdom.dom.element.meta.MetaUserDefinedElement;
import org.odftoolkit.simple.TextDocument;
import org.odftoolkit.simple.meta.Meta;


public class ODFOfficeDocumentHandler implements IOfficeDocumentHandler
{
	private static final String OFFICE_PROCESS = Messages.ODFOfficeDocumentHandler_OFFICEPATH; 
	
	// Praeferenzkey zum Pfad der Officeapplication
	//public static final String OFFICE_APPLICATION_PREF = "officeapplication_pref"; //$NON-NLS-1$
	
	protected File fileDoc;
	
	private TextDocument odfDocument;
	
	private static final String USERELEMENTNAME = "NtOffice"; 
	private String documentProperty;
	private String authorProperty = "Nt Office"; //$NON-NLS-N$

	@Override
	public void openOfficeDocument(File fileDoc)
	{
		odfDocument = null;

		String ext = FilenameUtils.getExtension(fileDoc.getPath());
		if (StringUtils.equals(ext, ODF_OFFICETEXTDOCUMENT_EXTENSION))
		{

			try
			{
				this.fileDoc = fileDoc;
				odfDocument = TextDocument.loadDocument(fileDoc);
				readDocumentProperty();

			} catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	private void readDocumentProperty()
	{
		documentProperty = null;
		Meta meta = odfDocument.getOfficeMetadata();
		MetaUserDefinedElement element = meta.getUserDefinedElementByAttributeName(USERELEMENTNAME);
		if(element != null)
			documentProperty = element.getTextContent();
	}
	
	@Override
	public void saveOfficeDocument()
	{
		try
		{			
			writeDocumentProperty();
			odfDocument.save(fileDoc);
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
	}

	private void writeDocumentProperty()
	{
		Meta meta = odfDocument.getOfficeMetadata();
		meta.setUserDefinedData(USERELEMENTNAME,"Text", documentProperty); //$NON-NLS-N$		
	}

	@Override
	public void showOfficeDocument()
	{
		//DrawDocument drawDocument;
	}
	
	//@Override
	/*
	public void showOfficeDocumentOLD()
	{
		
		Job j = new Job("Show Document") //$NON-NLS-1$
		{
			@Override
			protected IStatus run(IProgressMonitor monitor)
			{
				try
				{		
					// Linux 
					if(SystemUtils.IS_OS_LINUX)
					{
						Desktop.getDesktop().open(fileDoc);
						return Status.OK_STATUS;
					}
					
					// Pfad zur externen Officeanwendung (s. it.naturtalent.e4.office.ui.OfficeProcessor)
					String defPath = null;
					IEclipsePreferences defaultPreferenceNode = DefaultScope.INSTANCE
							.getNode(
									OfficeConstants.ROOT_OFFICE_PREFERENCES_NODE);
					// Pfad zur LibreOffice-Application aus dem DefaultPreference 
					if(defaultPreferenceNode != null)
						defPath = defaultPreferenceNode.get(OfficeConstants.OFFICE_APPLICATION_PREF, null);
						
					// Pfad zur LibreOffice-Application aus den Preferences
					IEclipsePreferences preferences = InstanceScope.INSTANCE
							  .getNode(OfficeConstants.ROOT_OFFICE_PREFERENCES_NODE);					 
					String unoPath = preferences.get(OfficeConstants.OFFICE_APPLICATION_PREF, defPath);
					if(StringUtils.isEmpty(unoPath))
						throw new java.io.IOException( Messages.ODFOfficeDocumentHandler_NO_OFFICE_ERROR );

					// Abbruch, wenn Pfad ungueltig
					File checkUnoPath = new File(unoPath);
					if(!checkUnoPath.exists() || !checkUnoPath.isDirectory())
					{						
						Display.getDefault().syncExec(new Runnable()
						{
							public void run()
							{
								MessageDialog
										.openError(
												Display.getDefault()
														.getActiveShell(),
												Messages.ODFOfficeDocumentHandler_OPEN_ERROR,
												Messages.ODFOfficeDocumentHandler_OFFICE_APPLICATION_PATH_ERROR);
							}
						});
						throw new java.io.IOException( Messages.ODFOfficeDocumentHandler_NO_OFFICE_ERROR );
					}

					// den Prozessaufruf vorbereiten
					unoPath = (new File(unoPath, OFFICE_PROCESS)).getPath();					
					String[] cmdArray;
					
					// Dokument schliessen (intern packen aller Komponenten)			
					if(fileDoc != null)
					{
						cmdArray = new String[2];
						cmdArray[0] = unoPath;
						cmdArray[1] = fileDoc.getPath();				
					}
					else
					{
						cmdArray = new String[1];
						cmdArray[0] = unoPath;
						//cmdArray[1] = "--writer";				
					}
					
					// externen Officeprozess starten
					Process process = Runtime.getRuntime().exec(cmdArray);
					if ( process == null )
						throw new Exception( Messages.ODFOfficeDocumentHandler_CANNOT_START_ERROR + cmdArray );
					
				} catch (Exception e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}		

				
				return Status.OK_STATUS;
			}
		};

		j.schedule();
	}
	*/
		
	@Override
	public void closeOfficeDocument()
	{
		if (odfDocument != null)
		{
			odfDocument.close();
			odfDocument = null;
		}
	}

	@Override
	public Object getDocument()
	{		
		return odfDocument;
	}

	@Override
	public String getProperty()
	{
		return documentProperty;
	}

	@Override
	public void setProperty(String property)
	{
		this.documentProperty = property;				
	}

	
		

}
