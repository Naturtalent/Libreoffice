package it.naturtalent.libreoffice.listeners;

import java.util.ArrayList;

import org.apache.commons.lang3.SystemUtils;

import com.sun.jna.Native;
import com.sun.jna.platform.win32.Kernel32;
import com.sun.jna.platform.win32.Tlhelp32;
import com.sun.jna.platform.win32.WinDef;
import com.sun.jna.platform.win32.WinNT;
import com.sun.jna.win32.W32APIOptions;

public class JNAClickListener
{
	//private Kernel32 kernel32;
	
	public static Kernel32 kernel32 = Kernel32.INSTANCE;
	
	public JNAClickListener()
	{
		if(SystemUtils.IS_OS_LINUX)
		{
			

			
			
		}
	}
	
	
	public ArrayList<Integer> getPIDs(String nm)
	// gather all the process IDs that starts with nm
	{
		final ArrayList<Integer> pids = new ArrayList<Integer>();

		WinNT.HANDLE snapshot = kernel32.CreateToolhelp32Snapshot(
				Tlhelp32.TH32CS_SNAPPROCESS, new WinDef.DWORD(0));
		Tlhelp32.PROCESSENTRY32.ByReference processEntry = new Tlhelp32.PROCESSENTRY32.ByReference();
		try
		{
			while (kernel32.Process32Next(snapshot, processEntry))
			{
				String processFnm = Native.toString(processEntry.szExeFile);
				if (processFnm.startsWith(nm))
				{
					int pid = processEntry.th32ProcessID.intValue();
					pids.add(pid);
					// int numThreads = processEntry.cntThreads.intValue();
					// System.out.println("Found " + nm + " (" + processFnm + ")
					// with pid=" + pid +
					// "; num threads: " + numThreads);
				}
			}
		} finally
		{
			kernel32.CloseHandle(snapshot);
		}
		return pids;
	} // end of getPIDs()
	
}
