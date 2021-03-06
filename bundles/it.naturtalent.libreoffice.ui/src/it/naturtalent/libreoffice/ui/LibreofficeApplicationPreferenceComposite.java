package it.naturtalent.libreoffice.ui;

import it.naturtalent.e4.preferences.DirectoryEditorComposite;
import it.naturtalent.libreoffice.Activator;
import it.naturtalent.libreoffice.utils.Lo;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.nio.file.FileSystems;
import java.nio.file.Paths;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.SystemUtils;
import org.eclipse.core.runtime.FileLocator;
import org.eclipse.core.runtime.Path;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Composite;
import org.osgi.framework.Bundle;
import org.osgi.framework.BundleContext;
import org.osgi.framework.FrameworkUtil;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;


/**
 * Composite der Libreoffice Preferaenzen
 * 
 * @author dieter
 *
 */
public class LibreofficeApplicationPreferenceComposite extends Composite
{

	private DirectoryEditorComposite directoryEditorComposite;
	
	private DirectoryEditorComposite jpipeDirectoryComposite;
	
	private DirectoryEditorComposite unoDirectoryComposite;
	
	// Templates der OS-spezifischen Skriptfiles   
	private static final String LINUX_KILLSCRIPT_NAME = "killScript.sh"; //$NON-NLS-1$	
	private static final String LINUX_KILLSCRIPT_PATH = "/templates/" + LINUX_KILLSCRIPT_NAME; //$NON-NLS-1$
	private static final String WINDOWS_KILLSCRIPT_NAME = "loKill.bat"; //$NON-NLS-1$
	private static final String WINDOWS_KILLSCRIPT_PATH = "/templates/" + WINDOWS_KILLSCRIPT_NAME; //$NON-NLS-1$
	
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
		
		Button btnTest = new Button(this, SWT.NONE);
		btnTest.addSelectionListener(new SelectionAdapter()
		{
			@Override
			public void widgetSelected(SelectionEvent e)
			{		
				start(directoryEditorComposite.getDirectory());
			}
		});
		btnTest.setBounds(15, 72, 132, 30);
		btnTest.setText("Start Libreoffice");
		
		
		// Button zum Abschiessen von Libreoffice
		Button btnKill = new Button(this, SWT.NONE);
		btnKill.setBounds(15, 126, 132, 30);
		btnKill.setText("Kill Libreoffice");
		btnKill.addSelectionListener(new SelectionAdapter()
		{
			@Override
			public void widgetSelected(SelectionEvent e)
			{
				String userDir = Paths.get(".").toAbsolutePath().normalize().toString();
				File scriptDir = new File(userDir);
								
				if (SystemUtils.IS_OS_LINUX)
				{					
					try
					{	
						// Kill-Skript in das aktuelle Verzeichnis "/home/..." kopieren
						copyScriptFile(scriptDir, LINUX_KILLSCRIPT_PATH);
						
						Runtime.getRuntime().exec("sh "+ LINUX_KILLSCRIPT_NAME);
						return;
						
					} catch (Exception e1)
					{
						System.out.println("Unable to kill Office: " + e1);
					}
				}
				
				if (SystemUtils.IS_OS_WINDOWS)
				{
					try
					{	
						// Kill-Skript in das aktuelle Verzeichnis "/home/..." kopieren
						copyScriptFile(scriptDir, WINDOWS_KILLSCRIPT_PATH);
						Runtime.getRuntime().exec("cmd /c "+WINDOWS_KILLSCRIPT_NAME);
						return;
						
					} catch (Exception e1)
					{
						System.out.println("Unable to kill Office: " + e1);
					}					
				}
			}
		});

		
		
	}
	
	private void start(String libreofficePath)
	{
		if (StringUtils.isNotEmpty(libreofficePath))
		{
			File libOfficeDir = new File(libreofficePath);
			if (libOfficeDir.isDirectory())
			{
				try
				{
					// via Runtime.exec 'libreoffice' starten
					File sofficeFile = new File(libOfficeDir, "soffice");
					Runtime.getRuntime().exec(new String[]{ sofficeFile.getPath() });
					
				} catch (IOException e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				return;
			}
		}
	}
	
	private void kill()
	{
		String userDir = Paths.get(".").toAbsolutePath().normalize().toString();
		File scriptDir = new File(userDir);
						
		if (SystemUtils.IS_OS_LINUX)
		{					
			try
			{	
				// Kill-Skript in das aktuelle Verzeichnis "/home/..." kopieren
				copyScriptFile(scriptDir, LINUX_KILLSCRIPT_PATH);
				
				Runtime.getRuntime().exec("sh "+ LINUX_KILLSCRIPT_NAME);
				return;
				
			} catch (Exception e1)
			{
				System.out.println("Unable to kill Office: " + e1);
			}
		}
		
		if (SystemUtils.IS_OS_WINDOWS)
		{
			try
			{	
				// Kill-Skript in das aktuelle Verzeichnis "/home/..." kopieren
				copyScriptFile(scriptDir, WINDOWS_KILLSCRIPT_PATH);
				Runtime.getRuntime().exec("cmd /c "+WINDOWS_KILLSCRIPT_NAME);
				return;
				
			} catch (Exception e1)
			{
				System.out.println("Unable to kill Office: " + e1);
			}					
		}
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

	
	/**
	 * Kopiert das Template des OS-spezifischen Kill-Skripts in das aktuelle Verzeichnis.
	 * 
	 * @param currentDir
	 * @param templateScriptPath
	 * @throws Exception
	 */
	private void copyScriptFile(File currentDir, String templateScriptPath) throws Exception
	{		
		File currentScriptFile = new File(currentDir, FilenameUtils.getName(templateScriptPath));
		if(!currentScriptFile.exists())
		{
			Bundle bundle = FrameworkUtil.getBundle(Activator.class);
			BundleContext bundleContext = bundle.getBundleContext();
			URL urlScript = FileLocator.find(bundleContext.getBundle(),new Path(templateScriptPath), null);
			urlScript = FileLocator.resolve(urlScript);
			try
			{				
				FileUtils.copyURLToFile(urlScript, currentScriptFile);				
			} catch (IOException e)
			{							
				//log.error(Messages.DesignUtils_ErrorCreateDrawFile);
				e.printStackTrace();
			}
		}	
	}


	@Override
	protected void checkSubclass()
	{
		// Disable the check that prevents subclassing of SWT components
	}
}
