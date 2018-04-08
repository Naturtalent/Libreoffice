package it.naturtalent.libreoffice.ui;

import it.naturtalent.e4.preferences.DirectoryEditorComposite;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Composite;

public class LibreofficeApplicationPreferenceComposite extends Composite
{

	private DirectoryEditorComposite directoryEditorComposite;
	
	private DirectoryEditorComposite jpipeDirectoryComposite;
	
	private DirectoryEditorComposite unoDirectoryComposite;
	
	/**
	 * Create the composite.
	 * @param parent
	 * @param style
	 */
	public LibreofficeApplicationPreferenceComposite(Composite parent, int style)
	{
		super(parent, style);
		setLayout(null);
		
		// Verzeichnis indem LibreOffice 'soffice' installiert ist
		directoryEditorComposite = new DirectoryEditorComposite(this, SWT.NONE);
		directoryEditorComposite.setBounds(5, 5, 529, 61);
				
		// Verzeichnis indem sich die 'JPipe' - Bibliothek befindet
		jpipeDirectoryComposite = new DirectoryEditorComposite(this, SWT.NONE);
		jpipeDirectoryComposite.setBounds(5, 99, 529, 61);
		jpipeDirectoryComposite.setLabel("Verzeichnis der JPIPE-Biblipthek");
		jpipeDirectoryComposite.setEnable(false);
		
		// Verzeichnis indem sich die UNO - Klassen befinden
		unoDirectoryComposite = new DirectoryEditorComposite(this, SWT.NONE);
		unoDirectoryComposite.setLabel("Verzeichnis LibreOffice UNO Klassen (juh.jar, jurt.jar...) ");
		unoDirectoryComposite.setBounds(10, 202, 529, 61);
		unoDirectoryComposite.setEnable(false);

	}
	
	

	public DirectoryEditorComposite getDirectoryEditorComposite()
	{
		return directoryEditorComposite;
	}
	
	public DirectoryEditorComposite getJpipeDirectoryComposite()
	{
		return jpipeDirectoryComposite;
	}

	public DirectoryEditorComposite getUnoDirectoryComposite()
	{
		return unoDirectoryComposite;
	}



	@Override
	protected void checkSubclass()
	{
		// Disable the check that prevents subclassing of SWT components
	}
}
