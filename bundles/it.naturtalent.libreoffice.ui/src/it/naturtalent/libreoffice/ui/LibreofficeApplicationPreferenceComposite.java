package it.naturtalent.libreoffice.ui;

import it.naturtalent.e4.preferences.DirectoryEditorComposite;
import it.naturtalent.libreoffice.utils.Lo;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;


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
		
		Button btnKillLibreoffice = new Button(this, SWT.NONE);
		btnKillLibreoffice.addSelectionListener(new SelectionAdapter()
		{
			@Override
			public void widgetSelected(SelectionEvent e)
			{
				Lo.killOffice();
				System.out.println("Kill Libreoffice");
			}
		});
		btnKillLibreoffice.setBounds(20, 290, 108, 25);
		btnKillLibreoffice.setText("Kill Libreoffice");

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
