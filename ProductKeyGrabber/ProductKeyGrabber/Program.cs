/*
 *   This program is free software: you can redistribute it and/or modify
 *   it under the terms of the GNU General Public License as published by
 *   the Free Software Foundation, either version 3 of the License, or
 *   (at your option) any later version.
 *
 *   This program is distributed in the hope that it will be useful,
 *   but WITHOUT ANY WARRANTY; without even the implied warranty of
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *   GNU General Public License for more details.
 *
 *   You should have received a copy of the GNU General Public License
 *   along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
 
using System;
using System.Collections;
using Microsoft.Win32;
using System.IO;
using System.Net.NetworkInformation;

public class Globals
{
	public static RegistryKey windowsRegistry;
	public static RegistryKey officeRegistry;
	public static ArrayList windowsKeys = new ArrayList();
	public static ArrayList officeKeys = new ArrayList();
	public static Ping ping = new Ping ();
	public static PingReply pingReply;
}

public class KeyKaptain
  {
	
    public static byte[] GetWindowsID(string hostname)
    {
    	byte[] digitalProductId = null;
    	try
    	{
    		// Note that this is a global handle and will be used for other needs
	    	Globals.windowsRegistry = 
	    		RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, hostname);
			Globals.windowsRegistry = 
				Globals.windowsRegistry.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
	    	if(Globals.windowsRegistry != null)
	    	{
	    		// TODO: For other products, key name may be different.
	        	digitalProductId = Globals.windowsRegistry.GetValue("DigitalProductId") as byte[];
	        	//Globals.windowsRegistry.Close();
	      	}
    	}
    	catch
    	{
    		digitalProductId = null;
    	}
      	return digitalProductId;         
    }
    
    public static byte[] GetOfficeID(string hostname)
    {
    	byte[] digitalProductId = null;
    	string[] valueNames = null;
		for (int i = 0 ; i < 25 ; i++) // Office versions currently go up to 12, but going up to 25 should be sufficient future-proofing
	    {
		   	try
		   	{
		   		Globals.officeRegistry = 
		   			RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, hostname); // we have no choice but to use multiple handles
			   	string temp = @"SOFTWARE\Microsoft\Office\" + i + @".0\Registration";
				Globals.officeRegistry = 
					Globals.officeRegistry.OpenSubKey(temp);
			   	if(Globals.officeRegistry != null)
			   	{
			   		valueNames = Globals.officeRegistry.GetSubKeyNames();
			   		Globals.officeRegistry =
			   			Globals.officeRegistry.OpenSubKey(valueNames[0]); // there should only be one
			   		try
			   		{
			   			digitalProductId = Globals.officeRegistry.GetValue("DigitalProductID") as byte[];
			   			break;
			   		}
			   		catch
			   		{
			   			// nothing needs to be done
			   		}
			 	}
		   	}
		   	catch
		   	{
		   		Globals.officeRegistry.Close();
		   	}
		}
      	return digitalProductId;         
    }
    
    public static string GetOfficeVersion()
    {
    	if(Globals.officeRegistry != null)
	    {
    		try
    		{
		        return Globals.officeRegistry.GetValue("ProductName") as string;
    		}
    		catch
    		{
    			return "";
    		}
	    }
    	return "";
    }
    
    public static string GetOwner()
    {
    	if(Globals.windowsRegistry != null)
	    {
    		try
    		{
		        return Globals.windowsRegistry.GetValue("RegisteredOwner") as string;
    		}
    		catch
    		{
    			return "";
    		}
	    }
    	return "";
    }
    
    public static string GetOS()
    {
    	if(Globals.windowsRegistry != null)
	    {
    		try
    		{
		        return Globals.windowsRegistry.GetValue("ProductName") as string;
    		}
    		catch
    		{
    			return "";
    		}
	    }
    	return "";
    }
    
    public static string GetOSVersion()
    {
    	if(Globals.windowsRegistry != null)
	    {
    		try
    		{
		        return Globals.windowsRegistry.GetValue("CurrentVersion") as string;
    		}
    		catch
    		{
    			return "";
    		}
	    }
    	return "";
    }
    
    public static string GetServicePack()
    {
    	if(Globals.windowsRegistry != null)
	    {
    		try
    		{
		        return Globals.windowsRegistry.GetValue("CSDVersion") as string;
    		}
    		catch
    		{
    			return "";
    		}
	    }
    	return "";
    }
    
    public static string GetProcessor(string hostname)
    {
    	string processor = null;
    	try
    	{
	    	RegistryKey reg = 
	    		RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, hostname);
			reg = 
				reg.OpenSubKey(@"HARDWARE\DESCRIPTION\System\CentralProcessor\0");
	    	if(reg != null)
	    	{
	        	processor = reg.GetValue("ProcessorNameString") as string;
	      	}
    	}
    	catch
    	{
    		processor = "";
    	}
      	return processor;         
    }
    
    public static string DecodeProductKey(byte[] digitalProductId)
    {
    	if (digitalProductId != null)
    	{
	      const int keyStartIndex = 52;
	      const int keyEndIndex = keyStartIndex + 15;
	      char[] digits = new char[]
	      {
	        'B', 'C', 'D', 'F', 'G', 'H', 'J', 'K', 'M', 'P', 'Q', 'R', 
	        'T', 'V', 'W', 'X', 'Y', '2', '3', '4', '6', '7', '8', '9',
	      };
	      const int decodeLength = 29;
	      const int decodeStringLength = 15;
	      char[] decodedChars = new char[decodeLength];
	      ArrayList hexPid = new ArrayList();
	      for (int i = keyStartIndex; i <= keyEndIndex; i++)
	      {
	        hexPid.Add(digitalProductId[i]);
	      }
	      for (int i = decodeLength - 1; i >= 0; i--)
	      {
	        if ((i + 1) % 6 == 0)
	        {
	          decodedChars[i] = '-';
	        }
	        else
	        {
	          int digitMapIndex = 0;
	          for (int j = decodeStringLength - 1; j >= 0; j--)
	          {
	            int byteValue = (digitMapIndex << 8) | (byte)hexPid[j];
	            hexPid[j] = (byte)(byteValue / 24);
	            digitMapIndex = byteValue % 24;
	            decodedChars[i] = digits[digitMapIndex];
	          }
	        }
	      }
	      return new string(decodedChars);
    	}
    	else
    	{
    		return "";
    	}
    }
    
    public static byte[] StringToByteArray(string str)
	{
	    System.Text.ASCIIEncoding  encoding = new System.Text.ASCIIEncoding();
	    return encoding.GetBytes(str);
	}
    
    public static int Main(string[] args)
    {
    	/*
    	 * Let me begin by saying I know the code below is messy, but it gets
    	 * the job done.  The code above is as close to procedural programming
    	 * as I can get, but I would still prefer something more C-like.  Global
    	 * variables, for example, are all wrong.
    	 * It would be great if someone would clean up the code below, but keep
    	 * in mind that people at ActiveForever use IE, as well as Firefox, so
    	 * keep it cross-browser.
    	 * 
    	 * The license variable was added in order for it to show up in the
    	 * compiled binary
    	 */
    	string license = @"
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.";

    	
    	StreamWriter writer = new StreamWriter("webpage.html");
    	StreamWriter windowsWriter = new StreamWriter("windowskeys.txt");
    	StreamWriter officeWriter = new StreamWriter("officekeys.txt");
    	int i = 0;
    	string color;
    	string webpage = @"
    	<html>
    	<head>
    		<title>Network Scan results</title>";
    	webpage += "<style type=\"text/css\">";
    	webpage += "body { margin: 0px; }";
    	webpage += "</style>";
    	webpage += "</head>";
    	webpage += "<body text=#0000A0>"; //onload=\"detectBrowser()\">";
    	webpage += @"<center>
    	<table width=2500 cellspacing=0 cellpadding=2 border=0 bordercolor=#000000>
    		<tr bgcolor=#3574EC7>
    			<td><font color=#FFFFFF>Hostname</font></td>
    			<td><font color=#FFFFFF>Windows CD Key</font></td>
    			<td><font color=#FFFFFF>Office CD Key</font></td>
    			<td><font color=#FFFFFF>Office version</font></td>
    			<td><font color=#FFFFFF>Owner</font></td>
    			<td><font color=#FFFFFF>OS</font></td>
    			<td><font color=#FFFFFF>OS version</font></td>
    			<td><font color=#FFFFFF>Service Pack</font></td>
    			<td><font color=#FFFFFF>Processor</font></td>
    		</tr>";
    	foreach (string hostname in args)
    	{
    		try // this is required because the ping class is useless
    		{
    			Globals.pingReply = Globals.ping.Send(hostname);
    			bool ponged = (Globals.pingReply.Status == IPStatus.Success ?
    			                true : false); // workaround for non-working C# if statements.  what a terrible compiler.
		    	if (ponged)
	    		{
		    		string windowsKey = DecodeProductKey(GetWindowsID(hostname));
		    		string officeKey = DecodeProductKey(GetOfficeID(hostname));
		    		if (windowsKey != "") // windowsKey would only be blank if the computer is not Windows
		    		{
			    		System.Console.WriteLine("Now scanning {0}", hostname);
			    		i++;
			    		if (i % 2 == 1)
			    		{
			    			color = "#FFFFFF";
			    		}
			    		else
			    		{
			    			color = "#F0F0F0";
			    		}
			    		webpage += "<tr bgcolor=" + color + "><td>" 
			    				+ hostname + "</td><td>"
			    				+ windowsKey
			    				+ "</td>\n<td>"
			    				+ officeKey
			    				+ "</td>\n<td>"
			    				+ GetOfficeVersion()
			    				+ "</td>\n<td>"
			    				+ GetOwner()
			    				+ "</td>\n<td>"
			    				+ GetOS()
			    				+ "</td>\n<td>"
			    				+ GetOSVersion()
			    				+ "</td>\n<td>"
			    				+ GetServicePack()
			    				+ "</td>\n<td>"
			    				+ GetProcessor(hostname)
			    				+ "</td>\n</tr>";
			    		Globals.windowsKeys.Add(windowsKey);
			    		Globals.officeKeys.Add(officeKey);
		    		}
	    		}
    		}
    		catch
    		{
    			// just passin' through...
    		}
	    }
	    webpage += @"
	    </table>
	    </center>
	    </body>";
	    writer.WriteLine(webpage);
	    writer.Close();
	    foreach(string key in Globals.windowsKeys)
	    {
	    	windowsWriter.WriteLine(key);
	    }
	    windowsWriter.Close();
	    foreach(string key in Globals.officeKeys)
	    {
	    	officeWriter.WriteLine(key);
	    }
	    officeWriter.Close();
	    Globals.windowsRegistry.Close();
	    Globals.officeRegistry.Close();
  		return 0;
    }
}