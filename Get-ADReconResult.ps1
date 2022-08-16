#requires -version 2
<#
.SYNOPSIS
  Script to import all ADRecon results into one PowerShell Custom Object and return it to the user.
.DESCRIPTION
  Does what it has to da
.PARAMETER PathToCSVFolder
    Path to the CSV folder with the AD Recon results
.INPUTS
  None
.OUTPUTS
  PowerShell object with all results to be stored in a variable
.NOTES
  Version:        1.0
  Author:         Michael Ritter (@bigmike)
  Creation Date:  <Date>
  Purpose/Change: Initial script development
  
.EXAMPLE
  $results = Get-ADReconResults
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Dot Source required Function Libraries
#. "C:\Scripts\Functions\Logging_Functions.ps1"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Log File Info
#$sLogPath = "C:\Windows\Temp"
#$sLogName = "<script_name>.log"
#$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

$sPathToCSVFolder=".\CSV-Files\"


### Variables for Spooler check
$sourceSpooler = @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;

namespace PingCastle.ExtractedCode
{
	public class rprn
	{
            [DllImport("Rpcrt4.dll", EntryPoint = "RpcBindingFromStringBindingW",
            CallingConvention = CallingConvention.StdCall,
            CharSet = CharSet.Unicode, SetLastError = false)]
            private static extern Int32 RpcBindingFromStringBinding(String bindingString, out IntPtr lpBinding);
            
            [DllImport("Rpcrt4.dll", EntryPoint = "NdrClientCall2", CallingConvention = CallingConvention.Cdecl,
                CharSet = CharSet.Unicode, SetLastError = false)]
            private static extern IntPtr NdrClientCall2x86(IntPtr pMIDL_STUB_DESC, IntPtr formatString, IntPtr args);
            
            [DllImport("Rpcrt4.dll", EntryPoint = "RpcBindingFree", CallingConvention = CallingConvention.StdCall,
                CharSet = CharSet.Unicode, SetLastError = false)]
            private static extern Int32 RpcBindingFree(ref IntPtr lpString);
            
            [DllImport("Rpcrt4.dll", EntryPoint = "RpcStringBindingComposeW", CallingConvention = CallingConvention.StdCall,
                CharSet = CharSet.Unicode, SetLastError = false)]
            private static extern Int32 RpcStringBindingCompose(
                String ObjUuid, String ProtSeq, String NetworkAddr, String Endpoint, String Options,
                out IntPtr lpBindingString
                );
                
            [DllImport("Rpcrt4.dll", EntryPoint = "RpcBindingSetOption", CallingConvention = CallingConvention.StdCall, SetLastError = false)]
            private static extern Int32 RpcBindingSetOption(IntPtr Binding, UInt32 Option, IntPtr OptionValue);

		[DllImport("Rpcrt4.dll", EntryPoint = "NdrClientCall2", CallingConvention = CallingConvention.Cdecl,
		   CharSet = CharSet.Unicode, SetLastError = false)]
		internal static extern IntPtr NdrClientCall2x64(IntPtr pMIDL_STUB_DESC, IntPtr formatString, ref IntPtr Handle);
        
        [DllImport("Rpcrt4.dll", EntryPoint = "NdrClientCall2", CallingConvention = CallingConvention.Cdecl,
			CharSet = CharSet.Unicode, SetLastError = false)]
		private static extern IntPtr NdrClientCall2x64(IntPtr intPtr1, IntPtr intPtr2, string pPrinterName, out IntPtr pHandle, string pDatatype, ref rprn.DEVMODE_CONTAINER pDevModeContainer, int AccessRequired);

		[DllImport("Rpcrt4.dll", EntryPoint = "NdrClientCall2", CallingConvention = CallingConvention.Cdecl,
			CharSet = CharSet.Unicode, SetLastError = false)]
		private static extern IntPtr NdrClientCall2x64(IntPtr intPtr1, IntPtr intPtr2, IntPtr hPrinter, uint fdwFlags, uint fdwOptions, string pszLocalMachine, uint dwPrinterLocal, IntPtr intPtr3);

		private static byte[] MIDL_ProcFormatStringx86 = new byte[] {
				0x00,0x48,0x00,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,
				0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x01,0x00,0x18,0x00,0x31,0x04,0x00,0x00,0x00,0x5c,0x08,0x00,0x40,0x00,0x46,0x06,0x08,0x05,
				0x00,0x00,0x01,0x00,0x00,0x00,0x0b,0x00,0x00,0x00,0x02,0x00,0x10,0x01,0x04,0x00,0x0a,0x00,0x0b,0x00,0x08,0x00,0x02,0x00,0x0b,0x01,0x0c,0x00,0x1e,
				0x00,0x48,0x00,0x10,0x00,0x08,0x00,0x70,0x00,0x14,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x02,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,
				0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x03,0x00,0x08,0x00,0x32,
				0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,
				0x04,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,
				0x48,0x00,0x00,0x00,0x00,0x05,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,
				0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x06,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,
				0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x07,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,
				0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x08,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,
				0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x09,0x00,0x08,0x00,
				0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,
				0x00,0x0a,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,
				0x00,0x48,0x00,0x00,0x00,0x00,0x0b,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,
				0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0c,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,
				0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0d,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,
				0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0e,0x00,0x08,0x00,0x32,0x00,0x00,0x00,
				0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0f,0x00,0x08,
				0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,
				0x00,0x00,0x10,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,
				0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x11,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,
				0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x12,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,
				0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x13,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,
				0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x14,0x00,0x08,0x00,0x32,0x00,0x00,
				0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x15,0x00,
				0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,
				0x00,0x00,0x00,0x16,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,
				0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x17,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,
				0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x18,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,
				0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x19,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,
				0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1a,0x00,0x08,0x00,0x32,0x00,
				0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1b,
				0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,
				0x00,0x00,0x00,0x00,0x1c,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,
				0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1d,0x00,0x08,0x00,0x30,0xe0,0x00,0x00,0x00,0x00,0x38,0x00,0x40,0x00,0x44,0x02,0x08,0x01,0x00,0x00,
				0x00,0x00,0x00,0x00,0x18,0x01,0x00,0x00,0x36,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1e,0x00,0x08,0x00,0x32,0x00,0x00,
				0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1f,0x00,
				0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,
				0x00,0x00,0x00,0x20,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,
				0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x21,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,
				0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x22,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,
				0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x23,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,
				0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x24,0x00,0x08,0x00,0x32,0x00,
				0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x25,
				0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x26,0x00,
				0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x27,0x00,0x08,
				0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,
				0x00,0x00,0x28,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,
				0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x29,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,
				0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2a,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,
				0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2b,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
				0x40,0x00,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2c,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,
				0x00,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2d,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,
				0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2e,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,
				0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2f,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,
				0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x30,0x00,0x08,0x00,0x32,
				0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,
				0x31,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x32,
				0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x33,0x00,
				0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,
				0x00,0x00,0x00,0x34,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,
				0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x35,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,
				0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x36,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x01,
				0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x37,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x01,0x00,
				0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x38,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,
				0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x39,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,
				0x00,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3a,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,
				0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3b,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,
				0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3c,0x00,0x08,0x00,
				0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,
				0x00,0x3d,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x04,0x00,0x08,0x00,
				0x00,0x48,0x00,0x00,0x00,0x00,0x3e,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x08,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x70,
				0x00,0x04,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3f,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x01,0x00,0x00,
				0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x40,0x00,0x04,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x01,0x00,0x00,0x00,
				0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x41,0x00,0x1c,0x00,0x30,0x40,0x00,0x00,0x00,0x00,0x3c,0x00,0x08,0x00,0x46,0x07,0x08,0x05,0x00,0x00,
				0x01,0x00,0x00,0x00,0x08,0x00,0x00,0x00,0x3a,0x00,0x48,0x00,0x04,0x00,0x08,0x00,0x48,0x00,0x08,0x00,0x08,0x00,0x0b,0x00,0x0c,0x00,0x02,0x00,0x48,
				0x00,0x10,0x00,0x08,0x00,0x0b,0x00,0x14,0x00,0x3e,0x00,0x70,0x00,0x18,0x00,0x08,0x00,0x00
            };

		private static byte[] MIDL_ProcFormatStringx64 = new byte[] {
				0x00,0x48,0x00,0x00,0x00,0x00,0x00,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
				0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x01,0x00,0x30,0x00,0x31,0x08,0x00,0x00,0x00,0x5c,0x08,0x00,0x40,0x00,0x46,0x06,
				0x0a,0x05,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x0b,0x00,0x00,0x00,0x02,0x00,0x10,0x01,0x08,0x00,0x0a,0x00,0x0b,0x00,0x10,0x00,0x02,0x00,0x0b,
				0x01,0x18,0x00,0x1e,0x00,0x48,0x00,0x20,0x00,0x08,0x00,0x70,0x00,0x28,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x02,0x00,0x10,0x00,0x32,0x00,
				0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,
				0x00,0x03,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,
				0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x04,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,
				0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x05,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,
				0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x06,0x00,0x10,0x00,0x32,0x00,0x00,
				0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,
				0x07,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,
				0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x08,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,
				0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x09,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,
				0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0a,0x00,0x10,0x00,0x32,0x00,0x00,0x00,
				0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0b,
				0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,
				0x00,0x48,0x00,0x00,0x00,0x00,0x0c,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
				0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0d,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,
				0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0e,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,
				0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x0f,0x00,
				0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,
				0x48,0x00,0x00,0x00,0x00,0x10,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
				0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x11,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,
				0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x12,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,
				0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x13,0x00,0x10,
				0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,
				0x00,0x00,0x00,0x00,0x14,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,
				0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x15,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,
				0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x16,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,
				0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x17,0x00,0x10,0x00,
				0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,
				0x00,0x00,0x00,0x18,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,
				0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x19,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,
				0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1a,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,
				0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1b,0x00,0x10,0x00,0x32,
				0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,
				0x00,0x00,0x1c,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,
				0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1d,0x00,0x10,0x00,0x30,0xe0,0x00,0x00,0x00,0x00,0x38,0x00,0x40,0x00,0x44,0x02,0x0a,0x01,0x00,0x00,
				0x00,0x00,0x00,0x00,0x00,0x00,0x18,0x01,0x00,0x00,0x32,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x1e,0x00,0x10,0x00,0x32,
				0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,
				0x00,0x00,0x1f,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,
				0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x20,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,
				0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x21,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,
				0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x22,0x00,0x10,0x00,0x32,0x00,
				0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,
				0x00,0x23,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,
				0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x24,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,
				0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x25,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,
				0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x26,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,
				0x00,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x27,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,
				0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x28,0x00,0x10,0x00,0x32,
				0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,
				0x00,0x00,0x29,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,
				0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2a,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,
				0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2b,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,
				0x00,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2c,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
				0x40,0x00,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2d,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,
				0x00,0x40,0x00,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2e,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,
				0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x2f,0x00,0x10,
				0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,
				0x00,0x00,0x00,0x00,0x30,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,
				0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x31,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x0a,0x01,0x00,0x00,
				0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x32,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x0a,0x01,0x00,
				0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x33,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,
				0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x34,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,
				0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x35,0x00,
				0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,
				0x48,0x00,0x00,0x00,0x00,0x36,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
				0x00,0x48,0x00,0x00,0x00,0x00,0x37,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
				0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x38,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,
				0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x39,0x00,0x08,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x0a,
				0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3a,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,
				0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3b,0x00,0x10,0x00,0x32,0x00,0x00,
				0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,
				0x3c,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,
				0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3d,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,
				0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3e,0x00,0x10,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x44,0x01,0x0a,
				0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x70,0x00,0x08,0x00,0x08,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x3f,0x00,0x08,0x00,0x32,0x00,0x00,0x00,
				0x00,0x00,0x00,0x00,0x40,0x00,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x40,0x00,0x08,0x00,0x32,0x00,0x00,
				0x00,0x00,0x00,0x00,0x00,0x40,0x00,0x0a,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x48,0x00,0x00,0x00,0x00,0x41,0x00,0x38,0x00,0x30,0x40,
				0x00,0x00,0x00,0x00,0x3c,0x00,0x08,0x00,0x46,0x07,0x0a,0x05,0x00,0x00,0x01,0x00,0x00,0x00,0x00,0x00,0x08,0x00,0x00,0x00,0x36,0x00,0x48,0x00,0x08,
				0x00,0x08,0x00,0x48,0x00,0x10,0x00,0x08,0x00,0x0b,0x00,0x18,0x00,0x02,0x00,0x48,0x00,0x20,0x00,0x08,0x00,0x0b,0x00,0x28,0x00,0x3a,0x00,0x70,0x00,
				0x30,0x00,0x08,0x00,0x00

        };

		private static byte[] MIDL_TypeFormatStringx86 = new byte[] {
				0x00,0x00,0x12,0x08,0x25,0x5c,0x11,0x04,0x02,0x00,0x30,0xa0,0x00,0x00,0x11,0x00,0x0e,0x00,0x1b,0x00,0x01,0x00,0x19,0x00,0x00,0x00,0x01,0x00,0x01,
				0x5b,0x16,0x03,0x08,0x00,0x4b,0x5c,0x46,0x5c,0x04,0x00,0x04,0x00,0x12,0x00,0xe6,0xff,0x5b,0x08,0x08,0x5b,0x11,0x04,0x02,0x00,0x30,0xe1,0x00,0x00,
				0x30,0x41,0x00,0x00,0x12,0x00,0x48,0x00,0x1b,0x01,0x02,0x00,0x19,0x00,0x0c,0x00,0x01,0x00,0x06,0x5b,0x16,0x03,0x14,0x00,0x4b,0x5c,0x46,0x5c,0x10,
				0x00,0x10,0x00,0x12,0x00,0xe6,0xff,0x5b,0x06,0x06,0x08,0x08,0x08,0x08,0x5b,0x1b,0x03,0x14,0x00,0x19,0x00,0x08,0x00,0x01,0x00,0x4b,0x5c,0x48,0x49,
				0x14,0x00,0x00,0x00,0x01,0x00,0x10,0x00,0x10,0x00,0x12,0x00,0xc2,0xff,0x5b,0x4c,0x00,0xc9,0xff,0x5b,0x16,0x03,0x10,0x00,0x4b,0x5c,0x46,0x5c,0x0c,
				0x00,0x0c,0x00,0x12,0x00,0xd0,0xff,0x5b,0x08,0x08,0x08,0x08,0x5b,0x00
        };

		private static byte[] MIDL_TypeFormatStringx64 = new byte[] {
				0x00,0x00,0x12,0x08,0x25,0x5c,0x11,0x04,0x02,0x00,0x30,0xa0,0x00,0x00,0x11,0x00,0x0e,0x00,0x1b,0x00,0x01,0x00,0x19,0x00,0x00,0x00,0x01,0x00,0x01,
				0x5b,0x1a,0x03,0x10,0x00,0x00,0x00,0x06,0x00,0x08,0x40,0x36,0x5b,0x12,0x00,0xe6,0xff,0x11,0x04,0x02,0x00,0x30,0xe1,0x00,0x00,0x30,0x41,0x00,0x00,
				0x12,0x00,0x38,0x00,0x1b,0x01,0x02,0x00,0x19,0x00,0x0c,0x00,0x01,0x00,0x06,0x5b,0x1a,0x03,0x18,0x00,0x00,0x00,0x0a,0x00,0x06,0x06,0x08,0x08,0x08,
				0x36,0x5c,0x5b,0x12,0x00,0xe2,0xff,0x21,0x03,0x00,0x00,0x19,0x00,0x08,0x00,0x01,0x00,0xff,0xff,0xff,0xff,0x00,0x00,0x4c,0x00,0xda,0xff,0x5c,0x5b,
				0x1a,0x03,0x18,0x00,0x00,0x00,0x08,0x00,0x08,0x08,0x08,0x40,0x36,0x5b,0x12,0x00,0xda,0xff,0x00
        };

		[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
		public rprn()
		{
			Guid interfaceId = new Guid("12345678-1234-ABCD-EF00-0123456789AB");
			if (IntPtr.Size == 8)
			{
				InitializeStub(interfaceId, MIDL_ProcFormatStringx64, MIDL_TypeFormatStringx64, "\\pipe\\spoolss", 1, 0);
			}
			else
			{
				InitializeStub(interfaceId, MIDL_ProcFormatStringx86, MIDL_TypeFormatStringx86, "\\pipe\\spoolss", 1, 0);
			}
		}

		[SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
		~rprn()
		{
			freeStub();
		}

		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
		public struct DEVMODE_CONTAINER
		{
			Int32 cbBuf;
			IntPtr pDevMode;
		}

		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
		public struct RPC_V2_NOTIFY_OPTIONS_TYPE
		{
			UInt16 Type;
			UInt16 Reserved0;
			UInt32 Reserved1;
			UInt32 Reserved2;
			UInt32 Count;
			IntPtr pFields;
		};

		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
		public struct RPC_V2_NOTIFY_OPTIONS
		{
			UInt32 Version;
			UInt32 Reserved;
			UInt32 Count;
			/* [unique][size_is] */
			RPC_V2_NOTIFY_OPTIONS_TYPE pTypes;
		};

		[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
		public Int32 RpcOpenPrinter(string pPrinterName, out IntPtr pHandle, string pDatatype, ref DEVMODE_CONTAINER pDevModeContainer, Int32 AccessRequired)
		{
			IntPtr result = IntPtr.Zero;
			IntPtr intptrPrinterName = Marshal.StringToHGlobalUni(pPrinterName);
			IntPtr intptrDatatype = Marshal.StringToHGlobalUni(pDatatype);
			pHandle = IntPtr.Zero;
			try
			{
				if (IntPtr.Size == 8)
				{
					result = NdrClientCall2x64(GetStubHandle(), GetProcStringHandle(36), pPrinterName, out pHandle, pDatatype, ref pDevModeContainer, AccessRequired);
				}
				else
				{
					IntPtr tempValue = IntPtr.Zero;
					GCHandle handle = GCHandle.Alloc(tempValue, GCHandleType.Pinned);
					IntPtr tempValuePointer = handle.AddrOfPinnedObject();
					GCHandle handleDevModeContainer = GCHandle.Alloc(pDevModeContainer, GCHandleType.Pinned);
					IntPtr tempValueDevModeContainer = handleDevModeContainer.AddrOfPinnedObject();
					try
					{
						result = CallNdrClientCall2x86(34, intptrPrinterName, tempValuePointer, intptrDatatype, tempValueDevModeContainer, new IntPtr(AccessRequired));
						// each pinvoke work on a copy of the arguments (without an out specifier)
						// get back the data
						pHandle = Marshal.ReadIntPtr(tempValuePointer);
					}
					finally
					{
						handle.Free();
						handleDevModeContainer.Free();
					}
				}
			}
			catch (SEHException)
			{
				Trace.WriteLine("RpcOpenPrinter failed 0x" + Marshal.GetExceptionCode().ToString("x"));
				return Marshal.GetExceptionCode();
			}
			finally
			{
				if (intptrPrinterName != IntPtr.Zero)
					Marshal.FreeHGlobal(intptrPrinterName);
				if (intptrDatatype != IntPtr.Zero)
					Marshal.FreeHGlobal(intptrDatatype);
			}
			return (int)result.ToInt64();
		}

		[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
		public Int32 RpcClosePrinter(ref IntPtr ServerHandle)
		{
			IntPtr result = IntPtr.Zero;
			try
			{
				if (IntPtr.Size == 8)
				{
					result = NdrClientCall2x64(GetStubHandle(), GetProcStringHandle(1076), ref ServerHandle);
				}
				else
				{
					IntPtr tempValue = ServerHandle;
					GCHandle handle = GCHandle.Alloc(tempValue, GCHandleType.Pinned);
					IntPtr tempValuePointer = handle.AddrOfPinnedObject();
					try
					{
						result = CallNdrClientCall2x86(1018, tempValuePointer);
						// each pinvoke work on a copy of the arguments (without an out specifier)
						// get back the data
						ServerHandle = Marshal.ReadIntPtr(tempValuePointer);
					}
					finally
					{
						handle.Free();
					}
				}
			}
			catch (SEHException)
			{
				Trace.WriteLine("RpcClosePrinter failed 0x" + Marshal.GetExceptionCode().ToString("x"));
				return Marshal.GetExceptionCode();
			}
			return (int)result.ToInt64();
		}

		[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
		public Int32 RpcRemoteFindFirstPrinterChangeNotificationEx(
			/* [in] */ IntPtr hPrinter,
			/* [in] */ UInt32 fdwFlags,
			/* [in] */ UInt32 fdwOptions,
			/* [unique][string][in] */ string pszLocalMachine,
			/* [in] */ UInt32 dwPrinterLocal)
		{
			IntPtr result = IntPtr.Zero;
			IntPtr intptrLocalMachine = Marshal.StringToHGlobalUni(pszLocalMachine);
			try
			{
				if (IntPtr.Size == 8)
				{
					result = NdrClientCall2x64(GetStubHandle(), GetProcStringHandle(2308), hPrinter, fdwFlags, fdwOptions, pszLocalMachine, dwPrinterLocal, IntPtr.Zero);
				}
				else
				{
					try
					{
						result = CallNdrClientCall2x86(2178, hPrinter, new IntPtr(fdwFlags), new IntPtr(fdwOptions), intptrLocalMachine, new IntPtr(dwPrinterLocal), IntPtr.Zero);
						// each pinvoke work on a copy of the arguments (without an out specifier)
						// get back the data
					}
					finally
					{
					}
				}
			}
			catch (SEHException)
			{
				Trace.WriteLine("RpcRemoteFindFirstPrinterChangeNotificationEx failed 0x" + Marshal.GetExceptionCode().ToString("x"));
				return Marshal.GetExceptionCode();
			}
			finally
			{
				if (intptrLocalMachine != IntPtr.Zero)
					Marshal.FreeHGlobal(intptrLocalMachine);
			}
			return (int)result.ToInt64();
		}

    
        private byte[] MIDL_ProcFormatString;
        private byte[] MIDL_TypeFormatString;
        private GCHandle procString;
        private GCHandle formatString;
        private GCHandle stub;
        private GCHandle faultoffsets;
        private GCHandle clientinterface;
        private GCHandle bindinghandle;
        private string PipeName;

        // important: keep a reference on delegate to avoid CallbackOnCollectedDelegate exception
        bind BindDelegate;
        unbind UnbindDelegate;
        allocmemory AllocateMemoryDelegate = AllocateMemory;
        freememory FreeMemoryDelegate = FreeMemory;

        // 5 seconds
        public UInt32 RPCTimeOut = 5000;

        [StructLayout(LayoutKind.Sequential)]
        private struct COMM_FAULT_OFFSETS
        {
            public short CommOffset;
            public short FaultOffset;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1049:TypesThatOwnNativeResourcesShouldBeDisposable"), StructLayout(LayoutKind.Sequential)]
        private struct GENERIC_BINDING_ROUTINE_PAIR
        {
            public IntPtr Bind;
            public IntPtr Unbind;
        }
        

        [StructLayout(LayoutKind.Sequential)]
        private struct RPC_VERSION
        {
            public ushort MajorVersion;
            public ushort MinorVersion;


            public static RPC_VERSION INTERFACE_VERSION = new RPC_VERSION(1, 0);
            public static RPC_VERSION SYNTAX_VERSION = new RPC_VERSION(2, 0);

            public RPC_VERSION(ushort InterfaceVersionMajor, ushort InterfaceVersionMinor)
            {
                MajorVersion = InterfaceVersionMajor;
                MinorVersion = InterfaceVersionMinor;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RPC_SYNTAX_IDENTIFIER
        {
            public Guid SyntaxGUID;
            public RPC_VERSION SyntaxVersion;
        }

        

        [StructLayout(LayoutKind.Sequential)]
        private struct RPC_CLIENT_INTERFACE
        {
            public uint Length;
            public RPC_SYNTAX_IDENTIFIER InterfaceId;
            public RPC_SYNTAX_IDENTIFIER TransferSyntax;
            public IntPtr /*PRPC_DISPATCH_TABLE*/ DispatchTable;
            public uint RpcProtseqEndpointCount;
            public IntPtr /*PRPC_PROTSEQ_ENDPOINT*/ RpcProtseqEndpoint;
            public IntPtr Reserved;
            public IntPtr InterpreterInfo;
            public uint Flags;

            public static Guid IID_SYNTAX = new Guid(0x8A885D04u, 0x1CEB, 0x11C9, 0x9F, 0xE8, 0x08, 0x00, 0x2B,
                                                              0x10,
                                                              0x48, 0x60);

            public RPC_CLIENT_INTERFACE(Guid iid, ushort InterfaceVersionMajor, ushort InterfaceVersionMinor)
            {
                Length = (uint)Marshal.SizeOf(typeof(RPC_CLIENT_INTERFACE));
                RPC_VERSION rpcVersion = new RPC_VERSION(InterfaceVersionMajor, InterfaceVersionMinor);
                InterfaceId = new RPC_SYNTAX_IDENTIFIER();
                InterfaceId.SyntaxGUID = iid;
                InterfaceId.SyntaxVersion = rpcVersion;
                rpcVersion = new RPC_VERSION(2, 0);
                TransferSyntax = new RPC_SYNTAX_IDENTIFIER();
                TransferSyntax.SyntaxGUID = IID_SYNTAX;
                TransferSyntax.SyntaxVersion = rpcVersion;
                DispatchTable = IntPtr.Zero;
                RpcProtseqEndpointCount = 0u;
                RpcProtseqEndpoint = IntPtr.Zero;
                Reserved = IntPtr.Zero;
                InterpreterInfo = IntPtr.Zero;
                Flags = 0u;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MIDL_STUB_DESC
        {
            public IntPtr /*RPC_CLIENT_INTERFACE*/ RpcInterfaceInformation;
            public IntPtr pfnAllocate;
            public IntPtr pfnFree;
            public IntPtr pAutoBindHandle;
            public IntPtr /*NDR_RUNDOWN*/ apfnNdrRundownRoutines;
            public IntPtr /*GENERIC_BINDING_ROUTINE_PAIR*/ aGenericBindingRoutinePairs;
            public IntPtr /*EXPR_EVAL*/ apfnExprEval;
            public IntPtr /*XMIT_ROUTINE_QUINTUPLE*/ aXmitQuintuple;
            public IntPtr pFormatTypes;
            public int fCheckBounds;
            /* Ndr library version. */
            public uint Version;
            public IntPtr /*MALLOC_FREE_STRUCT*/ pMallocFreeStruct;
            public int MIDLVersion;
            public IntPtr CommFaultOffsets;
            // New fields for version 3.0+
            public IntPtr /*USER_MARSHAL_ROUTINE_QUADRUPLE*/ aUserMarshalQuadruple;
            // Notify routines - added for NT5, MIDL 5.0
            public IntPtr /*NDR_NOTIFY_ROUTINE*/ NotifyRoutineTable;
            public IntPtr mFlags;
            // International support routines - added for 64bit post NT5
            public IntPtr /*NDR_CS_ROUTINES*/ CsRoutineTables;
            public IntPtr ProxyServerInfo;
            public IntPtr /*NDR_EXPR_DESC*/ pExprInfo;
            // Fields up to now present in win2000 release.

            public MIDL_STUB_DESC(IntPtr pFormatTypesPtr, IntPtr RpcInterfaceInformationPtr,
                                    IntPtr pfnAllocatePtr, IntPtr pfnFreePtr, IntPtr aGenericBindingRoutinePairsPtr)
            {
                pFormatTypes = pFormatTypesPtr;
                RpcInterfaceInformation = RpcInterfaceInformationPtr;
                CommFaultOffsets = IntPtr.Zero;
                pfnAllocate = pfnAllocatePtr;
                pfnFree = pfnFreePtr;
                pAutoBindHandle = IntPtr.Zero;
                apfnNdrRundownRoutines = IntPtr.Zero;
                aGenericBindingRoutinePairs = aGenericBindingRoutinePairsPtr;
                apfnExprEval = IntPtr.Zero;
                aXmitQuintuple = IntPtr.Zero;
                fCheckBounds = 1;
                Version = 0x50002u;
                pMallocFreeStruct = IntPtr.Zero;
                MIDLVersion = 0x8000253;
                aUserMarshalQuadruple = IntPtr.Zero;
                NotifyRoutineTable = IntPtr.Zero;
                mFlags = new IntPtr(0x00000001);
                CsRoutineTables = IntPtr.Zero;
                ProxyServerInfo = IntPtr.Zero;
                pExprInfo = IntPtr.Zero;
            }
        }

        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected void InitializeStub(Guid interfaceID, byte[] MIDL_ProcFormatString, byte[] MIDL_TypeFormatString, string pipe, ushort MajorVerson, ushort MinorVersion)
        {
            this.MIDL_ProcFormatString = MIDL_ProcFormatString;
            this.MIDL_TypeFormatString = MIDL_TypeFormatString;
            PipeName = pipe;
            procString = GCHandle.Alloc(this.MIDL_ProcFormatString, GCHandleType.Pinned);

            RPC_CLIENT_INTERFACE clientinterfaceObject = new RPC_CLIENT_INTERFACE(interfaceID, MajorVerson, MinorVersion);
            GENERIC_BINDING_ROUTINE_PAIR bindingObject = new GENERIC_BINDING_ROUTINE_PAIR();
            // important: keep a reference to avoid CallbakcOnCollectedDelegate Exception
            BindDelegate = Bind;
            UnbindDelegate = Unbind;
            bindingObject.Bind = Marshal.GetFunctionPointerForDelegate((bind)BindDelegate);
            bindingObject.Unbind = Marshal.GetFunctionPointerForDelegate((unbind)UnbindDelegate);

            COMM_FAULT_OFFSETS commFaultOffset = new COMM_FAULT_OFFSETS();
            commFaultOffset.CommOffset = -1;
            commFaultOffset.FaultOffset = -1;
            faultoffsets = GCHandle.Alloc(commFaultOffset, GCHandleType.Pinned);
            clientinterface = GCHandle.Alloc(clientinterfaceObject, GCHandleType.Pinned);
            formatString = GCHandle.Alloc(MIDL_TypeFormatString, GCHandleType.Pinned);
            bindinghandle = GCHandle.Alloc(bindingObject, GCHandleType.Pinned);

            MIDL_STUB_DESC stubObject = new MIDL_STUB_DESC(formatString.AddrOfPinnedObject(),
                                                            clientinterface.AddrOfPinnedObject(),
                                                            Marshal.GetFunctionPointerForDelegate(AllocateMemoryDelegate),
                                                            Marshal.GetFunctionPointerForDelegate(FreeMemoryDelegate),
                                                            bindinghandle.AddrOfPinnedObject());

            stub = GCHandle.Alloc(stubObject, GCHandleType.Pinned);
        }

        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected void freeStub()
        {
            procString.Free();
            faultoffsets.Free();
            clientinterface.Free();
            formatString.Free();
            bindinghandle.Free();
            stub.Free();
        }

        delegate IntPtr allocmemory(int size);
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected static IntPtr AllocateMemory(int size)
        {
            IntPtr memory = Marshal.AllocHGlobal(size);
            //Trace.WriteLine("allocating " + memory.ToString());
            return memory;
        }

        delegate void freememory(IntPtr memory);
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected static void FreeMemory(IntPtr memory)
        {
            //Trace.WriteLine("freeing " + memory.ToString());
            Marshal.FreeHGlobal(memory);
        }

        delegate IntPtr bind(IntPtr IntPtrserver);
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected IntPtr Bind (IntPtr IntPtrserver)
        {
            string server = Marshal.PtrToStringUni(IntPtrserver);
            IntPtr bindingstring = IntPtr.Zero;
            IntPtr binding = IntPtr.Zero;
            Int32 status;

            Trace.WriteLine("Binding to " + server + " " + PipeName);
            status = RpcStringBindingCompose(null, "ncacn_np", server, PipeName, null, out bindingstring);
            if (status != 0)
            {
                Trace.WriteLine("RpcStringBindingCompose failed with status 0x" + status.ToString("x"));
                return IntPtr.Zero;
            }
            status = RpcBindingFromStringBinding(Marshal.PtrToStringUni(bindingstring), out binding);
            RpcBindingFree(ref bindingstring);
            if (status != 0)
            {
                Trace.WriteLine("RpcBindingFromStringBinding failed with status 0x" + status.ToString("x"));
                return IntPtr.Zero;
            }

            status = RpcBindingSetOption(binding, 12, new IntPtr(RPCTimeOut));
            if (status != 0)
            {
                Trace.WriteLine("RpcBindingSetOption failed with status 0x" + status.ToString("x"));
            }
            Trace.WriteLine("binding ok (handle=" + binding + ")");
            return binding;
        }

        delegate void unbind(IntPtr IntPtrserver, IntPtr hBinding);
        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected static void Unbind(IntPtr IntPtrserver, IntPtr hBinding)
        {
            string server = Marshal.PtrToStringUni(IntPtrserver);
            Trace.WriteLine("unbinding " + server);
            RpcBindingFree(ref hBinding);
        }

        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected IntPtr GetProcStringHandle(int offset)
        {
            return Marshal.UnsafeAddrOfPinnedArrayElement(MIDL_ProcFormatString, offset);
        }

        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected IntPtr GetStubHandle()
        {
            return stub.AddrOfPinnedObject();
        }

        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected IntPtr CallNdrClientCall2x86(int offset, params IntPtr[] args)
        {

            GCHandle stackhandle = GCHandle.Alloc(args, GCHandleType.Pinned);
            IntPtr result;
            try
            {
                result = NdrClientCall2x86(GetStubHandle(), GetProcStringHandle(offset), stackhandle.AddrOfPinnedObject());
            }
            finally
            {
                stackhandle.Free();
            }
            return result;
        }
        
        public bool CheckIfTheSpoolerIsActive(string computer)
		{
			IntPtr hHandle = IntPtr.Zero;

			DEVMODE_CONTAINER devmodeContainer = new DEVMODE_CONTAINER();
			try
			{
				Int32 ret = RpcOpenPrinter("\\\\" + computer, out hHandle, null, ref devmodeContainer, 0);
				if (ret == 0)
				{
					return true;
				}
			}
			finally
			{
				if (hHandle != IntPtr.Zero)
					RpcClosePrinter(ref hHandle);
			}
			return false;
		}
    }

}

"@

Add-Type -TypeDefinition $sourceSpooler
$rprn = New-Object PingCastle.ExtractedCode.rprn

## END Variabled for Spooler check

#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Test-Port {
  [CmdletBinding()]
  param (
      [Parameter(ValueFromPipeline = $true, HelpMessage = 'Could be suffixed by :Port')]
      [String[]]$ComputerName,

      [Parameter(HelpMessage = 'Will be ignored if the port is given in the param ComputerName')]
      [Int]$Port = 5985,

      [Parameter(HelpMessage = 'Timeout in millisecond. Increase the value if you want to test Internet resources.')]
      [Int]$Timeout = 1000
  )

  begin {
      $result = [System.Collections.ArrayList]::new()
  }

  process {
      foreach ($originalComputerName in $ComputerName) {
          $remoteInfo = $originalComputerName.Split(":")
          if ($remoteInfo.count -eq 1) {
              # In case $ComputerName in the form of 'host'
              $remoteHostname = $originalComputerName
              $remotePort = $Port
          } elseif ($remoteInfo.count -eq 2) {
              # In case $ComputerName in the form of 'host:port',
              # we often get host and port to check in this form.
              $remoteHostname = $remoteInfo[0]
              $remotePort = $remoteInfo[1]
          } else {
              $msg = "Got unknown format for the parameter ComputerName: " `
                  + "[$originalComputerName]. " `
                  + "The allowed formats is [hostname] or [hostname:port]."
              Write-Error $msg
              return
          }

          $tcpClient = New-Object System.Net.Sockets.TcpClient
          $portOpened = $tcpClient.ConnectAsync($remoteHostname, $remotePort).Wait($Timeout)

          $null = $result.Add([PSCustomObject]@{
              RemoteHostname       = $remoteHostname
              RemotePort           = $remotePort
              PortOpened           = $portOpened
              TimeoutInMillisecond = $Timeout
              SourceHostname       = $env:COMPUTERNAME
              OriginalComputerName = $originalComputerName
              })
      }
  }

  end {
      return $result
  }
}

Function Get-SpoolStatus {
  Param(
      [parameter(Mandatory=$true,
      ValueFromPipeline=$true)]
      [String[]]
      $ComputerName
  )

  $ComputerName = $ComputerName.TrimEnd()
  $spoolstatus = $rprn.CheckIfTheSpoolerIsActive($ComputerName)

  if($spoolstatus) {
      Write-Host "Spooler on Domain Controller $ComputerName is active" -ForegroundColor Black -BackgroundColor Red
  } else {
      Write-Host "Spooler on Domain Controller $ComputerName is not active" -ForegroundColor Black -BackgroundColor Green
  }
}



Function Get-ADReconResults{
  Param()
  
  Begin{
    #Log-Write -LogPath $sLogFile -LineValue "<description of what is going on>..."
    # Import AD-Recon results into variables
    # Check if Users.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $UsersCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "Users.csv"
    if(!(Test-Path -Path $UsersCSVPath)){
      $UsersResults = @()
    } else{
      $UsersResults = Import-Csv -Path $UsersCSVPath
    }

    # Check if Groups.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $GroupsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "Groups.csv"
    if(!(Test-Path -Path $GroupsCSVPath)){
      $GroupsResults = @()
    } else{
      $GroupsResults = Import-Csv -Path $GroupsCSVPath
    }

    # Check if Computers.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $ComputersCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "Computers.csv"
    if(!(Test-Path -Path $ComputersCSVPath)){
      $ComputersResults = @()
    } else{
      $ComputersResults = Import-Csv -Path $ComputersCSVPath
    }

    # Check if ACLs.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $ACLsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "ACLs.csv"
    if(!(Test-Path -Path $ACLsCSVPath)){
      $ACLsResults = @()
    } else{
      $ACLsResults = Import-Csv -Path $ACLsCSVPath
    }

    # Check if GroupMembers.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $GroupMembersCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "GroupMembers.csv"
    if(!(Test-Path -Path $GroupMembersCSVPath)){
      $GroupMembersResults = @()
    } else{
      $GroupMembersResults = Import-Csv -Path $GroupMembersCSVPath
    }

    # Check if BitLockerRecoveryKeys.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $BitLockerRecoveryKeysCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "BitLockerRecoveryKeys.csv"
    if(!(Test-Path -Path $BitLockerRecoveryKeysCSVPath)){
      $BitLockerRecoveryKeysResults = @()
    } else{
      $BitLockerRecoveryKeysResults = Import-Csv -Path $BitLockerRecoveryKeysCSVPath
    }

    # Check if ComputerSPNs.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $ComputerSPNsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "ComputerSPNs.csv"
    if(!(Test-Path -Path $ComputerSPNsCSVPath)){
      $ComputerSPNsResults = @()
    } else{
      $ComputerSPNsResults = Import-Csv -Path $ComputerSPNsCSVPath
    }

    # Check if DefaultPasswordPolicy.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $DefaultPasswordPolicyCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "DefaultPasswordPolicy.csv"
    if(!(Test-Path -Path $DefaultPasswordPolicyCSVPath)){
      $DefaultPasswordPolicyResults = @()
    } else{
      $DefaultPasswordPolicyResults = Import-Csv -Path $DefaultPasswordPolicyCSVPath
    }

    # Check if DNSNodes.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $DNSNodesCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "DNSNodes.csv"
    if(!(Test-Path -Path $DNSNodesCSVPath)){
      $DNSNodesResults = @()
    } else{
      $DNSNodesResults = Import-Csv -Path $DNSNodesCSVPath
    }

    # Check if DNSZones.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $DNSZonesCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "DNSZones.csv"
    if(!(Test-Path -Path $DNSZonesCSVPath)){
      $DNSZonesResults = @()
    } else{
      $DNSZonesResults = Import-Csv -Path $DNSZonesCSVPath
    }

    # Check if Domain.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $DomainCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "Domain.csv"
    if(!(Test-Path -Path $DomainCSVPath)){
      $DomainResults = @()
    } else{
      $DomainResults = Import-Csv -Path $DomainCSVPath
    }

    # Check if DomainControllers.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $DomainControllersCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "DomainControllers.csv"
    if(!(Test-Path -Path $DomainControllersCSVPath)){
      $DomainControllersResults = @()
    } else{
      $DomainControllersResults = Import-Csv -Path $DomainControllersCSVPath
    }

    # Check if FineGrainedPasswordPolicy.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $FineGrainedPasswordPolicyCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "FineGrainedPasswordPolicy.csv"
    if(!(Test-Path -Path $FineGrainedPasswordPolicyCSVPath)){
      $FineGrainedPasswordPolicyResults = @()
    } else{
      $FineGrainedPasswordPolicyResults = Import-Csv -Path $FineGrainedPasswordPolicyCSVPath
    }

    # Check if Forest.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $ForestCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "Forest.csv"
    if(!(Test-Path -Path $ForestCSVPath)){
      $ForestResults = @()
    } else{
      $ForestResults = Import-Csv -Path $ForestCSVPath
    }

    # Check if GPOs.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $GPOsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "GPOs.csv"
    if(!(Test-Path -Path $GPOsCSVPath)){
      $GPOsResults = @()
    } else{
      $GPOsResults = Import-Csv -Path $GPOsCSVPath
    }

    # Check if LAPS.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $LAPsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "LAPS.csv"
    if(!(Test-Path -Path $LAPsCSVPath)){
      $LAPsResults = @()
    } else{
      $LAPsResults = Import-Csv -Path $LAPsCSVPath
    }

    # Check if OUs.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $OUsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "OUs.csv"
    if(!(Test-Path -Path $OUsCSVPath)){
      $OUsResults = @()
    } else{
      $OUsResults = Import-Csv -Path $OUsCSVPath
    }

    # Check if Sites.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $SitesCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "Sites.csv"
    if(!(Test-Path -Path $SitesCSVPath)){
      $SitesResults = @()
    } else{
      $SitesResults = Import-Csv -Path $SitesCSVPath
    }

    # Check if Subnets.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $SubnetsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "Subnets.csv"
    if(!(Test-Path -Path $SubnetsCSVPath)){
      $SubnetsResults = @()
    } else{
      $SubnetsResults = Import-Csv -Path $SubnetsCSVPath
    }

    # Check if Trusts.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $TrustsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "Trusts.csv"
    if(!(Test-Path -Path $TrustsCSVPath)){
      $TrustsResults = @()
    } else{
      $TrustsResults = Import-Csv -Path $TrustsCSVPath
    }

    # Check if UserSPNs.csv file exists
    # If not, return empty array
    # If so, import results into array
    # Return array
    $UserSPNsCSVPath = Join-Path -Path $sPathToCSVFolder -ChildPath "UserSPNs.csv"
    if(!(Test-Path -Path $UserSPNsCSVPath)){
      $UserSPNsResults = @()
    } else{
      $UserSPNsResults = Import-Csv -Path $UserSPNsCSVPath
    }
  }
  
  Process{

      $results = [PSCustomObject]@{
                              ACLs     = $ACLsResults
                              BitLockerRecoveryKeys = $BitLockerRecoveryKeysResults
                              Computers = $ComputersResults
                              ComputerSPNs = $ComputerSPNsResults
                              DefaultPasswordPolicy = $DefaultPasswordPolicyResults
                              DNSNodes = $DNSNodesResults
                              DNSZones = $DNSZonesResults
                              Domain = $DomainResults
                              DomainControllers = $DomainControllersResults
                              FineGrainedPasswordPolicy = $FineGrainedPasswordPolicyResults
                              Forest = $ForestResults
                              GPOs = $GPOsResults
                              Groups = $GroupsResults
                              GroupMembers = $GroupMembersResults
                              LAPS = $LAPsResults
                              OUs = $OUsResults
                              Sites = $SitesResults
                              Subnets = $SubnetsResults
                              Trusts = $TrustsResults
                              Users = $UsersResults
                              UserSPNs = $UserSPNsResults
                              }
  }
  
  End{
    return $results
    }
  }
# write a function to calculate the date difference between two dates in days
function Get-DifferenceBetweenDates($startDate, $endDate) {
  $startDate = [datetime]::Parse($startDate)
  $endDate = [datetime]::Parse($endDate)
  $difference = $endDate - $startDate
  $difference = $difference.TotalDays
  return $difference
}

function Add-SecurityCheckItem {
  <#
  .SYNOPSIS
  Creates a new security check item and adds it to a global array.
  Author: Michael Ritter
  License: BSD 3-Clause
  .DESCRIPTION
  Single Security Checks that cannot be exported as CSV need to be collected centrally
  .PARAMETER SecurityItem
  Specifies the desired name for the security item group. (i.e. Microsoft PowerShell)
  .PARAMETER SecurityItemCheck
  Specifies the desired name for the specific check
  .PARAMETER AuditCheckResult
  Specifies the result of the check
  .PARAMETER AuditCheckPass
  Specifies if the security check was successful or not
  .EXAMPLE
  $result = "PowerShell v$($($PSVersionTable.PSVersion).Major) is installed and starts by default, important security features are shipped with this version" 
  Add-SecurityCheckItem -SecurityItem "PowerShell Version" -SecurityItemCheck "Check if at least PowerShell version 5 is in use" -AuditCheckResult $result -AuditCheckPass $true -FindingText $FindingText
  
  #>
  param (
      [Parameter(Position = 0, Mandatory=$True)]
      [String] $SecurityItem,
      [Parameter(Position = 1, Mandatory=$True)]
      [String] $SecurityItemCheck,
      [Parameter(Position = 2, Mandatory=$True)]
      [String] $AuditCheckResult,
      [Parameter(Position = 3, Mandatory=$True)]
      [Bool] $AuditCheckPass,
      [Parameter(Position = 3, Mandatory=$True)]
      [String] $FindingText
  )
      $SecurityItemAuditResults = @()
      $auditDetails = @{
          SecurityItem    = $SecurityItem
          Check     = $SecurityItemCheck
          Result      = $AuditCheckResult
          Passed = $AuditCheckPass
          FindingText= $FindingText
      } 
  
     [array]$Global:SecurityItemAuditResults += New-Object PSObject -Property $auditDetails
}


# MAIN
$results = Get-ADReconResults

# current directory is the path to the Reporting folder
$ReportingPath = "$((Get-Location).Path)/Reporting"

# check if the Reporting folder exists
if(!(Test-Path -Path $ReportingPath)){
  New-Item -Path $ReportingPath -ItemType Directory
}

# Stats
$TotalNumberOfComputers = $results.Computers.Count
$TotalNumberOfUsers = $results.Users.Count
$TotalNumberOfGroups = $results.Groups.Count
$TotalNumberOfTrusts = $results.Trusts.Count
$TotalNumberOfDomainControllers = $results.DomainControllers.Count
$TotalNumberOfDomain = $results.DomainControllers.Count
$TotalNumberOfLAPS = $results.LAPS.Count
$TotalNumberOfGPOs = $results.GPOs.Count
$TotalNumberOfBitLockerRecoveryKeys = $results.BitLockerRecoveryKeys.Count
$TotalNumberOfComputerSPNs = $results.ComputerSPNs.Count
$TotalNumberOfUserSPNs = $results.UserSPNs.Count
$TotalNumberOfGroupsMembers = $results.GroupMembers.Count
$AdminUsersCount = ($results.Users | Select-Object UserName, Enabled, AdminCount | Where-Object {$_.AdminCount -eq 1}).Count
$DomainName=($results.Domain | Where-Object {$_.Category -eq "Name"}).Value
$DomainSID=($results.Domain | Where-Object {$_.Category -eq "DomainSID"}).Value
$DomainFunctionalLevel=($results.Domain | Where-Object {$_.Category -eq "Functional Level"}).Value
$ForestName=($results.Forest | Where-Object {$_.Category -eq "Name"}).Value
$ForestFunctionalLevel=($results.Forest | Where-Object {$_.Category -eq "Functional Level"}).Value
$ForestDomains=($results.Forest | Where-Object {$_.Category -eq "Domain"}).Value

#-----------------------------------------------------------[Findings]------------------------------------------------------------
Write-Host '#################################################' -BackgroundColor Black
Write-Host '##              Forest Overview                ##' -BackgroundColor Black
Write-Host '#################################################' -BackgroundColor Black
Write-Host "Forest Name: $ForestName" -ForegroundColor Black -BackgroundColor White
Write-Host "Existing domains in the forest:" -ForegroundColor Black -BackgroundColor White
$ForestDomains

Write-Host '#################################################' -BackgroundColor Black
Write-Host '##              Forest Functional Level        ##' -BackgroundColor Black
Write-Host '#################################################' -BackgroundColor Black
Write-Host "Forest Name: $ForestName" -ForegroundColor Black -BackgroundColor White
Write-Host "Checking Forest functional level" -ForegroundColor Black -BackgroundColor White
if($ForestFunctionalLevel -match "2016"){
  Write-Host "Forest functional level is Windows Server 2016. This is a good level for a domain" -ForegroundColor Green -BackgroundColor White
} elseif($ForestFunctionalLevel -match "2008"){
  Write-Host "Forest functional level is Windows Server 2008. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($ForestFunctionalLevel -match "2003"){
  Write-Host "Forest functional level is Windows Server 2003. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($ForestFunctionalLevel -match "2000"){
  Write-Host "Forest functional level is Windows Server 2000.  The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($ForestFunctionalLevel -match "2003R2"){
  Write-Host "Forest functional level is Windows Server 2003R2. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($ForestFunctionalLevel -match "2008R2"){
  Write-Host "Forest functional level is Windows Server 2008R2. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($ForestFunctionalLevel -match "2012R2"){
  Write-Host "Forest functional level is Windows Server 2012R2. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} else{
  Write-Host "Forest functional level is unknown, please check manually" -ForegroundColor Black -BackgroundColor Yellow
}
Write-Host "For more information please refer to https://docs.microsoft.com/en-us/windows-server/identity/ad-ds/active-directory-functional-levels" -ForegroundColor Black -BackgroundColor White

Write-Host '#################################################' -BackgroundColor Black
Write-Host '##              Domain Functional Level        ##' -BackgroundColor Black
Write-Host '#################################################' -BackgroundColor Black
Write-Host "Domain Name: $domainName"
Write-Host "Checking Domain functional level" -ForegroundColor Black -BackgroundColor White
if($domainFunctionalLevel -match "2016"){
  Write-Host "Domain functional level is Windows Server 2016. This is a good level for a domain" -ForegroundColor Green -BackgroundColor White
} elseif($domainFunctionalLevel -match "2008"){
  Write-Host "Domain functional level is Windows Server 2008. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($domainFunctionalLevel -match "2003"){
  Write-Host "Domain functional level is Windows Server 2003. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($domainFunctionalLevel -match "2000"){
  Write-Host "Domain functional level is Windows Server 2000.  The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($domainFunctionalLevel -match "2003R2"){
  Write-Host "Domain functional level is Windows Server 2003R2. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($domainFunctionalLevel -match "2008R2"){
  Write-Host "Domain functional level is Windows Server 2008R2. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} elseif($domainFunctionalLevel -match "2012R2"){
  Write-Host "Domain functional level is Windows Server 2012R2. The domain does not support many security features. Functional Level should be updated to at least 2016 or higher" -ForegroundColor Black -BackgroundColor Red
} else{
  Write-Host "Domain functional level is unknown, please check manually" -ForegroundColor Black -BackgroundColor Yellow
}
Write-Host "For more information please refer to https://docs.microsoft.com/en-us/windows-server/identity/ad-ds/active-directory-functional-levels" -ForegroundColor Black -BackgroundColor White

# Read Passwords from GPOs (GPPPreferences and AutoLogon)
# Trusts
Write-Host '#################################################' -BackgroundColor Black
Write-Host '##          Domain Controllers with SMBv1      ##' -BackgroundColor Black
Write-Host '#################################################' -BackgroundColor Black
Write-Host "Domain Name: $domainName"
Write-Host "Checking Domain Domain Controllers for SMBv1 support" -ForegroundColor Black -BackgroundColor White
$DCwithSMBv1 = $results.DomainControllers | Select-Object Hostname, 
                                                          "SMB Port Open", 
                                                          "SMB1(NT LM 0.12)" | 
                                            Where-Object {$_."SMB1(NT LM 0.12)" -eq $true}

if($DCwithSMBv1.Count -eq 0){
  Write-Host "No Domain Controllers with SMBv1 support found" -ForegroundColor Black -BackgroundColor Green
} else {
  Write-Host "Domain Controllers with SMBv1 support found" -ForegroundColor Black -BackgroundColor Red
  $DCwithSMBv1 | Select-Object Hostname, "SMB Port Open", "SMB1(NT LM 0.12)" | Format-Table -AutoSize
}

Write-Host '##########################################################' -BackgroundColor Black
Write-Host '##          Domain Controllers without SMB-Signing      ##' -BackgroundColor Black
Write-Host '##########################################################' -BackgroundColor Black
Write-Host "Domain Name: $domainName"
Write-Host "Checking Domain Domain Controllers for SMB signing enforcement" -ForegroundColor Black -BackgroundColor White
$DCswithoutSMBSigning = $results.DomainControllers | Select-Object Hostname, 
                                                          "SMB Port Open", 
                                                          "SMB Signing" | 
                                            Where-Object {$_."SMB Signing" -eq $false}

if($DCswithoutSMBSigning.Count -gt 0){
  Write-Host "Domain Controllers with missing enforcement for SMB Signing identified" -ForegroundColor Black -BackgroundColor Red
  $DCswithoutSMBSigning | Select-Object Hostname, "SMB Port Open", "SMB Signing" | Format-Table -AutoSize
} else {
  Write-Host "No Domain Controllers idenified that do not enforce SMB Signing" -ForegroundColor Black -BackgroundColor Green
}

# Trusts
Write-Host '#########################################' -BackgroundColor Black
Write-Host '##              Trusts                 ##' -BackgroundColor Black
Write-Host '#########################################' -BackgroundColor Black
Write-Host 'Checking for trusted domains and the trust attributes' -ForegroundColor Black -BackgroundColor White
$Trusts = $results.Trusts
$TrustsWithinForest= $Trusts | Where-Object {$_.'Attributes' -eq "Within Forest"}
$TrustsOutsideForest = $Trusts | Where-Object {$_.'Attributes' -ne "Within Forest"}
Write-Host "There are $($Trusts.'Target Domain'.Count) trusts in from $domainName to other domains" -ForegroundColor Black -BackgroundColor DarkGray
Write-Host "There are $($TrustsWithinForest.'Target Domain'.Count) trusts within the forest" -ForegroundColor Black -BackgroundColor Yellow
$TrustsWithinForest | Format-Table -AutoSize
Write-Host "There are $($TrustsOutsideForest.'Target Domain'.Count) trusts from $domainName to other domains in externals forests" -ForegroundColor Black -BackgroundColor Red
$TrustsOutsideForest | Format-Table -AutoSize

# Unconstrained Delegation (User/Computers)
## Unconstrained Delegation Users:
Write-Host '#################################################################' -BackgroundColor Black
Write-Host '##              Unconstrained/Contrained Delegation            ##' -BackgroundColor Black
Write-Host '#################################################################' -BackgroundColor Black
Write-Host 'Checking for computers/users with delegration attributes defined' -ForegroundColor Black -BackgroundColor White

$UsersWithDelegration = $results.Users | Where-Object {-not ([string]::IsNullOrEmpty($_.'Delegation Type'))} |
                 Select-Object UserName, Enabled, 'Delegation Type'
$ComputersWithDelegation = $results.Computers | Where-Object {-not ([string]::IsNullOrEmpty($_.'Delegation Type'))} |
                  Select-Object DNSHostName, Enabled, 'Delegation Type'

if($UsersWithDelegation.Count -gt 0){
  Write-Host "There are $(($UsersWithDelegation.Count).ToString()) users with delegation enabled within the $domainName domain" -BackgroundColor Black -ForegroundColor Red
  $UsersWithDelegration | Format-Table -AutoSize
}
else
{
  Write-Host "There are $(($UsersWithDelegation.Count).ToString()) users with delegation enabled within the $domainName domain" -BackgroundColor Black -ForegroundColor Green
}

if($ComputersWithDelegation.Count -gt 0){
  Write-Host "There are $(($ComputersWithDelegation.Count).ToString()) computers with delegation enabled within the $domainName domain" -BackgroundColor Black -ForegroundColor Red
  $ComputersWithDelegation | Format-Table -AutoSize
}else{
  Write-Host "There are $(($ComputersWithDelegation.Count).ToString()) computers with delegation enabled within the $domainName domain" -BackgroundColor Black -ForegroundColor Green
}




# Insecure Password Policy in use
## Check Default Domain Policy for Password Settings
## Check if there is a seperate Domain Password Policy
## Check for existing fine grained password policies
Write-Host '#############################################' -BackgroundColor Black
Write-Host '##              Password Policy            ##' -BackgroundColor Black
Write-Host '#############################################' -BackgroundColor Black
Write-Host "Checking for Default Password Policy implementation of the $domainName Domain. Please review manually" -ForegroundColor Black -BackgroundColor White
$results.DefaultPasswordPolicy | Format-Table Policy,'Current Value','CIS Benchmark 2016' -AutoSize

# Kerberoasting
## Kerberoast it all and crack the hashes
## Kerberoast only accounts that are older than x years
## Execute targeted Kerberoasting if possible
Write-Host '#############################################' -BackgroundColor Black
Write-Host '##              Kerberoasting              ##' -BackgroundColor Black
Write-Host '#############################################' -BackgroundColor Black
Write-Host 'Checking for Kerberoastable Accounts' -ForegroundColor Black -BackgroundColor White
if($results.UserSPNs.Count -gt 0){
  Write-Host "There are $(($results.UserSPNs).count) Kerberoastable accounts in the $domainName domain... Happy Roasting" -ForegroundColor Black -BackgroundColor Red
  $results.UserSPNs | Format-Table -AutoSize

}else{
  Write-Host "There are no Kerberoastable accounts in the $domainName domain... No Kerberoasts today" -ForegroundColor Black -BackgroundColor Green
}

# Kerberoasting
## Kerberoast it all and crack the hashes
## Kerberoast only accounts that are older than x years
## Execute targeted Kerberoasting if possible
Write-Host '#############################################' -BackgroundColor Black
Write-Host '##              ASREP Roasting             ##' -BackgroundColor Black
Write-Host '#############################################' -BackgroundColor Black
Write-Host 'Checking for ASREPRoastable Accounts' -ForegroundColor Black -BackgroundColor White
$ASREPRoastableUsers = $results.Users | Where-Object {$_.'Does not Require Pre Auth' -eq $true}

if($ASREPRoastableUsers.Count -gt 0)
{
  Write-Host "There are $(($ASREPRoastableUsers.Count).ToString()) ASREPRoastable accounts in the $domainName domain... Happy Roasting" -ForegroundColor Black -BackgroundColor Red
  $ASREPRoastableUsers | Format-Table -AutoSize
}
else{
  Write-Host "There are $(($ASREPRoastableUsers.Count).ToString()) ASREPRoastable accounts in the $domainName domain... No ASREPRoast today" -ForegroundColor Black -BackgroundColor Green
}


# Missing rotation of krbtgt Password
## Check when the password of the krbtgt account was set the last time
Write-Host '#############################################' -BackgroundColor Black
Write-Host '##              Krbtgt Account              ##' -BackgroundColor Black
Write-Host '#############################################' -BackgroundColor Black
Write-Host "Checking the Password rotation of the krbtgt account within the $domainName domain" -ForegroundColor Black -BackgroundColor White
$krbtgtAccount=$results.Users | Select-Object UserName, 'Password Age (days)', 'Password LastSet' | 
                 Where-Object {$_.UserName -eq 'krbtgt'}

if([int]$krbtgtAccount.'Password Age (days)' -lt 90)
{
  Write-Host "The krbtgt account password was last set less than 90 days ago on $(($krbtgtAccount.'Password LastSet'.ToString()))" -ForegroundColor Black -BackgroundColor Green
}
elseif([int]$krbtgtAccount.'Password Age (days)' -lt 120){
  Write-Host "The krbtgt account password was last set on on $(($krbtgtAccount.'Password LastSet'.ToString())) less than 120 days" -ForegroundColor Black -BackgroundColor Yellow
}
elseif([int]$krbtgtAccount.'Password Age (days)' -lt 365){
  Write-Host "The krbtgt account password was last set on $(($krbtgtAccount.'Password LastSet'.ToString())) less than 365 days" -ForegroundColor Black -BackgroundColor Orange
}
else{
  Write-Host "The krbtgt account password was last set $(($krbtgtAccount.'Password Age (days)').ToString())) days ago. Password should be rotated regulary" -ForegroundColor Black -BackgroundColor Red
}
# Cleartext passwords in LDAP Attributes
## Check if the description attribute of user, computer or something else contains cleartext passwords
Write-Host '################################################' -BackgroundColor Black
Write-Host '##       Sensitive data in LDAP Attibutes     ##' -BackgroundColor Black
Write-Host '################################################' -BackgroundColor Black
Write-Host "Checking, if we can find sensitive data in ldap attributes..." -ForegroundColor Black -BackgroundColor White
$hitwords = @("pass",
              "pw",
              "secret",
              "abrechnung",
              "strom",
              "kalender",
              "nis")

$UsersWithSensitiveDataInAttributes=@()

ForEach($hitword in $hitwords)
{
  $identifiedUsers=$results.Description | Where-Object {$_.UserName -match "$hitword"}
  $UsersWithSensitiveDataInAttributes+=$identifiedUsers
}


# Protected Users
## Check if there are protected users
## make sure the sensitive accounts in the domain are in the protected users group
Write-Host '#############################################' -BackgroundColor Black
Write-Host '##       Protected Users Group             ##' -BackgroundColor Black
Write-Host '#############################################' -BackgroundColor Black
Write-Host "Checking the use of the protected users group within the $domainName domain" -ForegroundColor Black -BackgroundColor White
$ProtectedUsers = $results.GroupMembers | Where-Object {$_.'Group Name' -eq 'Protected Users'}
$ProtectedUsersCount=$($results.GroupMembers | Where-Object {$_.'Group Name' -eq 'Protected Users'}).'Member UserName'.Count

if($ProtectedUsersCount -gt 0)
{
  if($adminUsersCount -gt $ProtectedUsersCount)
  {
    Write-Host "There are $($ProtectedUsersCount) Protected Users and $($adminUsersCount) Admin Users in the $domainName domain. Admin Users should be in protected users group" -ForegroundColor Black -BackgroundColor Yellow
    $ProtectedUsers | Format-Table -AutoSize
    $ProtectedUsers | Export-Csv -FilePath "$ReportingPath/UsersInProtectedUsersGroup.csv" -NoTypeInformation -Encoding utf8
  }
  else{
    Write-Host "There are $($ProtectedUsersCount) Protected Users and $($adminUsersCount) Admin Users in the $domainName domain. All Admin users are in the protected users group" -ForegroundColor Black -BackgroundColor Green
    $ProtectedUsers | Format-Table -AutoSize
  }
}
else{
  Write-Host "There are no protected users in the $domainName domain" -ForegroundColor Black -BackgroundColor Red
}

# Cannot be deletgated Flag
## Make sure all privileged accounts are marked as cannot be delegated
# Check all privileged groups recursively beforehand and then for each user check if the cannot be delegated flag is set
Write-Host '#################################################' -BackgroundColor Black
Write-Host '##         Cannot be delegated flag            ##' -BackgroundColor Black
Write-Host '#################################################' -BackgroundColor Black
Write-Host 'Checking for users that cannot be delegated' -ForegroundColor Black -BackgroundColor White
$UsersThatCannnotBeDelegated = $results.Users | Select-Object UserName, 'Enabled', 'Delegation Permitted' |  Where-Object {$_.'Delegation Permitted' -eq $false} 

$UsersThatCannotBeDelegatedCount = $($UsersThatCannnotBeDelegated).Count
$UsersThatCannotBeDelegatedCountPercentage = $([math]::Round($UsersThatCannotBeDelegatedCount/$TotalNumberOfUsers*100,2)).ToString()

Write-Host "Number of users that cannot be delegated: $UsersThatCannotBeDelegatedCount"
Write-Host "Percentage of users that cannot be delegated: $UsersThatCannotBeDelegatedCountPercentage%"

if($UsersThatCannotBeDelegatedCount -gt 0)
{
  if($adminUsersCount -gt $UsersThatCannotBeDelegatedCount)
  {
    Write-Host "There are $($UsersThatCannotBeDelegatedCount) users that cannot be delegated and $($adminUsersCount) admin users. Admin Users should have the cannot be delegated flag set" -ForegroundColor Black -BackgroundColor Yellow
    $UsersThatCannnotBeDelegated | Format-Table UserName, Enabled, 'Delegation Permitted' -AutoSize
  }
  else{
    Write-Host "There are $($UsersThatCannotBeDelegatedCount) users that cannot be delegated and $($adminUsersCount) admin users. All admin users have the cannot be delegated flag set" -ForegroundColor Black -BackgroundColor Green
    $UsersThatCannnotBeDelegated | Format-Table UserName, Enabled, 'Delegation Permitted' -AutoSize
  }
}
else{
  Write-Host "There are no users that have the cannot be delegated flag set" -ForegroundColor Black -BackgroundColor Red
}

# Print Spooler
## Check if the print spooler is enabled on the domain controllers
## https://github.com/carlospolop/hacktricks/blob/master/windows/active-directory-methodology/printers-spooler-service-abuse.md
# 
Write-Host '#################################################' -BackgroundColor Black
Write-Host '##                  PrinterSpooler             ##' -BackgroundColor Black
Write-Host '#################################################' -BackgroundColor Black
Write-Host 'Please execute the printer spooler test for each domain controller manually' -ForegroundColor Black -BackgroundColor White
Write-Host 'Script: https://github.com/NotMedic/NetNTLMtoSilverTicket/blob/master/Get-SpoolStatus.ps1' -ForegroundColor Black -BackgroundColor Yellow
foreach($hostname in $results.DomainControllers.Hostname)
{
  Write-Host "Get-SpoolStatus $hostname"
}

Write-Host '#################################################' -BackgroundColor Black
Write-Host '##                  Never logged in            ##' -BackgroundColor Black
Write-Host '#################################################' -BackgroundColor Black
Write-Host 'Checking for users that have never signed in' -ForegroundColor Black -BackgroundColor White
# Never signed in
## Check if the never signed in flag is set
$UsersNeverSignedIn = $results.Users | Select-Object UserName, Enabled, 'Never Logged in' | 
                                       Where-Object {$_.'Never Logged in' -eq $true}

$UsersNeverSignedInCount = $($UsersNeverSignedIn.Count)

if($UsersNeverSignedInCount -gt 0)
{
    $UsersNeverSignedInEnabled = $UsersNeverSignedIn | Where-Object {$_.'Enabled' -eq $true}
    $UsersNeverSignedInDisabled = $UsersNeverSignedIn | Where-Object {$_.'Enabled' -eq $false}

    $UsersNeverSignedInEnabledCount = $($UsersNeverSignedInEnabled.Count)
    $UsersNeverSignedInDisabledCount = $($UsersNeverSignedInDisabled.Count)

    $UsersNeverSignedInPercentage = $([math]::Round($UsersNeverSignedInCount/$TotalNumberOfUsers*100,2))
    $UsersNeverSignedInEnabledPercentage = $([math]::Round($UsersNeverSignedInEnabledCount/$UsersNeverSignedInCount*100,2))
    $UsersNeverSignedInDisabledPercentage = $([math]::Round($UsersNeverSignedInDisabledCount/$UsersNeverSignedInCount*100,2))


    Write-Host "Total users that never signed in: $UsersNeverSignedInCount" -BackgroundColor DarkGray -ForegroundColor Black
    Write-Host "Percentage users that never signed in: $UsersNeverSignedInPercentage%" -BackgroundColor DarkGray -ForegroundColor Black
    Write-Host "Users that never signed in (Enabled): $UsersNeverSignedInEnabledCount" -BackgroundColor Red -ForegroundColor Black
    Write-Host "Percentage of users that never signed in (Enabled): $UsersNeverSignedInEnabledPercentage%" -BackgroundColor Red -ForegroundColor Black
    Write-Host "Users that never signed in (Disabled): $UsersNeverSignedInDisabledCount" -BackgroundColor Yellow -ForegroundColor Black
    Write-Host "Users that never signed in (Disabled) (%) : $UsersNeverSignedInDisabledPercentage%" -BackgroundColor Yellow -ForegroundColor Black
    
    $FindingText_DE="Während unserer Tests haben wir festgestellt, dass sich in der Domäne $domainName, welche insgesamt $([string]::Format('{0:N0}',$TotalNumberOfUsers)) Benutzerkonten verwaltet, $([string]::Format('{0:N0}',$UsersNeverSignedInCount)) ($UsersNeverSignedInPercentage%) Benutzerkonten befinden, welche sich noch nie am Active Directory angemeldet haben. Davon sind aktuell $([string]::Format('{0:N0}',$UsersNeverSignedInEnabledCount)) ($UsersNeverSignedInEnabledPercentage%) Benutzerkonten als aktiv gelistet und $([string]::Format('{0:N0}',$UsersNeverSignedInDisabledCount)) ($UsersNeverSignedInDisabledPercentage%) deaktiviert."
    $FindingText_EN="During our tests we noticed, that in the domain $domainName, which manages a total $([string]::Format('{0:N0}',$TotalNumberOfUsers)) user accounts, $([string]::Format('{0:N0}',$UsersNeverSignedInCount)) ($UsersNeverSignedInPercentage%) accounts have never logged in the Active Directory. Of these, currently $([string]::Format('{0:N0}',$UsersNeverSignedInEnabledCount)) ($UsersNeverSignedInEnabledPercentage%) are marked as active and $([string]::Format('{0:N0}',$UsersNeverSignedInDisabledCount)) ($UsersNeverSignedInDisabledPercentage%) inactive."

    Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor White
    Write-Host $FindingText_EN -ForegroundColor Black -BackgroundColor White

    $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-text-users-never-signed-in.txt" -Encoding utf8
    "`r`n" | Out-File -FilePath "$ReportingPath/finding-text-users-never-signed-in.txt" -Append -Encoding utf8
    $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-text-users-never-signed-in.txt" -Append -Encoding utf8

    $UsersNeverSignedIn | Export-Csv -FilePath "$ReportingPath/UsersNeverSignedIn.csv" -NoTypeInformation -Encoding utf8
}
else{
  Write-Host "There are no users that have never signed in" -ForegroundColor Black -BackgroundColor Green
}


# SeMachineAccountPrivilege
## Check if users are able to add computers to the domain


Write-Host '#################################################' -BackgroundColor Black
Write-Host '##          Password change next login         ##' -BackgroundColor Black
Write-Host '#################################################' -BackgroundColor Black
Write-Host 'Checking for users that have to change the password on the next login' -ForegroundColor Black -BackgroundColor White
# Accounts to be prompted to change Password
## Check if there are accounts that need to be prompted to change their password
$UsersMustChangePwdNextLogon = $results.Users | Select-Object UserName, Enabled, 'Must Change Password at Logon' | 
                                              Where-Object {$_.'Must Change Password at Logon' -eq $true}


$UsersMustChangePwdNextLogonCount = $($UsersMustChangePwdNextLogon.Count)
if($UsersMustChangePwdNextLogonCount -gt 0)
{
    $UsersMustChangePwdNextLogonEnabled = $UsersMustChangePwdNextLogon | Where-Object {$_.'Enabled' -eq $true}
    $UsersMustChangePwdNextLogonDisabled = $UsersMustChangePwdNextLogon | Where-Object {$_.'Enabled' -eq $false}

    $UsersMustChangePwdNextLogonEnabledCount = $($UsersMustChangePwdNextLogonEnabled.UserName.count).ToString()
    $UsersMustChangePwdNextLogonDisabledCount = $($UsersMustChangePwdNextLogonDisabled.UserName.Count).ToString()

    $UsersMustChangePwdNextLogonPercentage = $([math]::Round($UsersMustChangePwdNextLogonCount/$TotalNumberOfUsers*100,2))
    $UsersMustChangePwdNextLogonEnabledPercentage = $([math]::Round($UsersMustChangePwdNextLogonEnabledCount/$UsersMustChangePwdNextLogonCount*100,2))
    $UsersMustChangePwdNextLogonDisabledPercentage = $([math]::Round($UsersMustChangePwdNextLogonDisabledCount/$UsersMustChangePwdNextLogonCount*100,2))

    Write-Host "Total users with that must change password next logon: $UsersMustChangePwdNextLogonCount" -BackgroundColor DarkGray -ForegroundColor Black
    Write-Host "Total users with that must change password next logon (%): $UsersMustChangePwdNextLogonPercentage%" -BackgroundColor DarkGray -ForegroundColor Black
    Write-Host "Users with that must change password next logon (Enabled): $UsersMustChangePwdNextLogonEnabledCount" -BackgroundColor Red -ForegroundColor Black
    Write-Host "Users with that must change password next logon (Enabled) (%) : $UsersMustChangePwdNextLogonEnabledPercentage%" -BackgroundColor Red -ForegroundColor Black
    Write-Host "Users with that must change password next logon (Disabled): $UsersMustChangePwdNextLogonDisabledCount" -BackgroundColor Yellow -ForegroundColor Black
    Write-Host "Users with that must change password next logon (Disabled) (%) : $UsersMustChangePwdNextLogonDisabledPercentage%" -BackgroundColor Yellow -ForegroundColor Black

    $FindingText_DE="Während unserer Tests haben wir festgestellt, dass sich in der Domäne $domainName, welche insgesamt $([string]::Format('{0:N0}',$TotalNumberOfUsers)) Benutzerkonten verwaltet, $([string]::Format('{0:N0}',$UsersMustChangePwdNextLogonCount)) ($UsersMustChangePwdNextLogonPercentage%) Benutzerkonten befinden, welche ihr Passwort bei der nächsten Anmeldung aktualisieren müssen. Davon sind aktuell $([string]::Format('{0:N0}',$UsersMustChangePwdNextLogonEnabledCount)) ($UsersMustChangePwdNextLogonEnabledPercentage%) Benutzerkonten als aktiv gelistet und $([string]::Format('{0:N0}',$UsersMustChangePwdNextLogonDisabledCount)) ($UsersMustChangePwdNextLogonDisabledPercentage%) deaktiviert."
    $FindingText_EN="During our testing, we found that in the $domainName domain, which manages a total of $([string]::Format('{0:N0}',$TotalNumberOfUsers)) user accounts, there are $([string]::Format('{0:N0}',$UsersMustChangePwdNextLogonCount)) ($UsersMustChangePwdNextLogonPercentage%) user accounts that need to update their password at the next login. Of these, $([string]::Format('{0:N0}',$UsersMustChangePwdNextLogonEnabledCount)) ($(UsersMustChangePwdNextLogonEnabledPercentage)%) are currently listed as active and $([string]::Format('{0:N0}',$UsersMustChangePwdNextLogonDisabledCount)) ($UsersMustChangePwdNextLogonDisabledPercentage%) are disabled."
  
    Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor White
    Write-Host $FindingText_EN -ForegroundColor Black -BackgroundColor White

    $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-text-must-change-pass-next-logon.txt" -Encoding utf8
    "`r`n" | Out-File -FilePath "$ReportingPath/finding-text-must-change-pass-next-logon.txt" -Append -Encoding utf8
    $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-text-must-change-pass-next-logon.txt" -Append -Encoding utf8

    $UsersMustChangePwdNextLogon | Export-Csv -Path "$ReportingPath/UserMustChangePasswdNextLogon.csv" -NoTypeInformation -Encoding UTF8
}
else{
  Write-Host "There are no users that need to change the password on the next login" -ForegroundColor Black -BackgroundColor Red
}


Write-Host '##########################################################' -BackgroundColor Black
Write-Host '##                  No password expiry attribute        ##' -BackgroundColor Black
Write-Host '##########################################################' -BackgroundColor Black
Write-Host 'Checking for users that have no password expiry' -ForegroundColor Black -BackgroundColor White
# User accounts without password expiry
## Check if there are accounts that have no password expiry set
$UsersWithNoPasswordExpiry = $results.Users | Select-Object UserName, Enabled, 'Password Never Expires' | 
                                              Where-Object {$_.'Password Never Expires' -eq $true}

$UsersWithNoPasswordExpiryCount = $($UsersWithNoPasswordExpiry.Count)

if($UsersWithNoPasswordExpiryCount -gt 0){
  
  $UsersWithNoPasswordExpiryEnabled = $UsersWithNoPasswordExpiry | Where-Object {$_.'Enabled' -eq $true}
  $UsersWithNoPasswordExpiryDisabled = $UsersWithNoPasswordExpiry | Where-Object {$_.'Enabled' -eq $false}
  
  $UsersWithNoPasswordExpiryEnabledCount = $($UsersWithNoPasswordExpiryEnabled.Count).ToString()
  $UsersWithNoPasswordExpiryDisabledCount = $($UsersWithNoPasswordExpiryDisabled.Count).ToString()
  
  $UsersWithNoPasswordExpiryPercentage = $([math]::Round($UsersWithNoPasswordExpiryCount/$TotalNumberOfUsers*100,2))
  $UsersWithNoPasswordExpiryEnabledPercentage = $([math]::Round($UsersWithNoPasswordExpiryEnabledCount/$UsersWithNoPasswordExpiryCount*100,2))
  $UsersWithNoPasswordExpiryDisabledPercentage = $([math]::Round($UsersWithNoPasswordExpiryDisabledCount/$UsersWithNoPasswordExpiryCount*100,2))
  
  Write-Host "Total users with no password expiry: $UsersWithNoPasswordExpiryCount" -BackgroundColor DarkGray -ForegroundColor Black
  Write-Host "Total users with no password expiry (%): $UsersWithNoPasswordExpiryPercentage%" -BackgroundColor DarkGray -ForegroundColor Black
  Write-Host "Total users with no password expiry (Enabled): $UsersWithNoPasswordExpiryEnabledCount" -BackgroundColor Red -ForegroundColor Black
  Write-Host "Total users with no password expiry (Enabled) (%) : $UsersWithNoPasswordExpiryEnabledPercentage%" -BackgroundColor Red -ForegroundColor Black
  Write-Host "Total users with no password expiry (Disabled): $UsersWithNoPasswordExpiryDisabledCount" -BackgroundColor Yellow -ForegroundColor Black
  Write-Host "Total users with no password expiry (Disabled) (%) : $UsersWithNoPasswordExpiryDisabledPercentage%" -BackgroundColor Yellow -ForegroundColor Black

  $FindingText_DE="Während unserer Tests haben wir festgestellt, dass sich in der Domäne $domainName, welche insgesamt $([string]::Format('{0:N0}',$TotalNumberOfUsers)) Benutzerkonten verwaltet, $([string]::Format('{0:N0}',$UsersWithNoPasswordExpiryCount)) ($UsersWithNoPasswordExpiryPercentage%) Benutzerkonten befinden, bei denen das Passwort niemals abläuft und erneut werden muss. Davon sind aktuell $([string]::Format('{0:N0}',$UsersWithNoPasswordExpiryEnabledCount)) ($UsersWithNoPasswordExpiryEnabledPercentage%) Benutzerkonten als aktiv gelistet und $([string]::Format('{0:N0}',$UsersWithNoPasswordExpiryDisabledCount)) ($UsersWithNoPasswordExpiryDisabledPercentage%) deaktiviert."
  $FindingText_EN="During our testing, we found that in the $domainName domain, which manages a total of $([string]::Format('{0:N0}',$TotalNumberOfUsers)) user accounts, there are $([string]::Format('{0:N0}',$UsersWithNoPasswordExpiryCount)) ($UsersWithNoPasswordExpiryPercentage%) user accounts where the password never expires and needs to be reissued. Of these, $([string]::Format('{0:N0}',$UsersWithNoPasswordExpiryEnabledCount)) ($UsersWithNoPasswordExpiryEnabledPercentage%) user accounts are currently listed as active and $([string]::Format('{0:N0}',$UsersWithNoPasswordExpiryDisabledCount)) ($UsersWithNoPasswordExpiryDisabledPercentage%) are disabled."


  Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor White
  Write-Host $FindingText_EN -ForegroundColor Black -BackgroundColor White

  $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-no-password-expiry.txt" -Encoding utf8
  "`r`n" | Out-File -FilePath "$ReportingPath/finding-no-password-expiry.txt" -Append -Encoding utf8
  $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-no-password-expiry.txt" -Append -Encoding utf8

  $UsersWithNoPasswordExpiry | Export-Csv -Path "$ReportingPath/UsersWithNoPasswordExpiry.csv" -NoTypeInformation -Encoding UTF8
}
else{
  Write-Host "There are no users that have the no password expiry attribute set" -ForegroundColor Black -BackgroundColor Green
}


Write-Host '#######################################################' -BackgroundColor Black
Write-Host '##               Password Age < 365 days             ##' -BackgroundColor Black
Write-Host '#######################################################' -BackgroundColor Black
Write-Host 'Checking for users with password age < 365 days' -ForegroundColor Black -BackgroundColor White
# User Accounts with password age < 365 days
## Check if there are accounts that have a password age > 365 days
$UsersWithBigPasswordAge = $results.Users | Select-Object UserName, 
                                                Enabled, 
                                                @{Name="PasswordAge_days";Expression={[int] $_.'Password Age (days)'}} | 
                                            Where-Object {($_.PasswordAge_days -gt 90)}

$UsersWithBigPasswordAgeCount = $($UsersWithBigPasswordAge.Count)

If($UsersWithBigPasswordAgeCount -gt 0)
{
      $UsersWithBigPasswordAgeEnabled = $UsersWithBigPasswordAge | Where-Object {$_.Enabled -eq $true}
      $UsersWithBigPasswordAgeDisabled = $UsersWithBigPasswordAge | Where-Object {$_.Enabled -eq $false}
      
      $UsersWithBigPasswordAgeEnabledCount = $($UsersWithBigPasswordAgeEnabled.Count)
      $UsersWithBigPasswordAgeDisabledCount = $($UsersWithBigPasswordAgeDisabled.Count)
      
      $UsersWithBigPasswordAgePercentage = $([math]::Round($UsersWithBigPasswordAgeCount/$TotalNumberOfUsers*100,2))
      $UsersWithBigPasswordAgeEnabledPercentage = $([math]::Round($UsersWithBigPasswordAgeEnabledCount/$UsersWithBigPasswordAgeCount*100,2))
      $UsersWithBigPasswordAgeDisabledPercentage = $([math]::Round($UsersWithBigPasswordAgeDisabledCount/$UsersWithBigPasswordAgeCount*100,2))
      
      Write-Host "Total users with a password age > 365 days: $UsersWithBigPasswordAgeCount" -BackgroundColor DarkGray -ForegroundColor Black
      Write-Host "Total users with a password age > 365 days: $UsersWithBigPasswordAgePercentage%" -BackgroundColor DarkGray -ForegroundColor Black
      Write-Host "Users with a password age > 365 days (Enabled): $UsersWithBigPasswordAgeEnabledCount" -BackgroundColor Red -ForegroundColor Black
      Write-Host "Users with a password age > 365 days (Enabled): $UsersWithBigPasswordAgeEnabledPercentage%" -BackgroundColor Red -ForegroundColor Black
      Write-Host "Users with a password age > 365 days (Disabled): $UsersWithBigPasswordAgeDisabledCount" -BackgroundColor Yellow -ForegroundColor Black
      Write-Host "Users with a password age > 365 days (Disabled): $UsersWithBigPasswordAgeDisabledPercentage%" -BackgroundColor Yellow -ForegroundColor Black

      $FindingText_DE="Während unserer Tests haben wir festgestellt, dass sich in der Domäne $domainName, welche insgesamt $([string]::Format('{0:N0}',$TotalNumberOfUsers)) Benutzerkonten verwaltet, $([string]::Format('{0:N0}',$UsersWithBigPasswordAgeCount)) ($UsersWithBigPasswordAgePercentage%) Benutzerkonten befinden, bei denen das Passwort älter als 365 Tage ist. Davon sind aktuell $([string]::Format('{0:N0}',$UsersWithBigPasswordAgeEnabledCount)) ($UsersWithBigPasswordAgeEnabledPercentage%) Benutzerkonten als aktiv gelistet und $([string]::Format('{0:N0}',$UsersWithBigPasswordAgeDisabledCount)) ($UsersWithBigPasswordAgeDisabledPercentage%) deaktiviert."
      $FindingText_EN="During our tests, we found that in the $domainName domain, which manages a total of $([string]::Format('{0:N0}',$TotalNumberOfUsers)) user accounts, there are $([string]::Format('{0:N0}',$UsersWithBigPasswordAgeCount)) ($UsersWithBigPasswordAgePercentage%) user accounts where the password is older than 365 days. Of these, $([string]::Format('{0:N0}',$UsersWithBigPasswordAgeEnabledCount)) ($UsersWithBigPasswordAgeEnabledPercentage%) user accounts are currently listed as active and $([string]::Format('{0:N0}',$UsersWithBigPasswordAgeDisabledCount)) ($UsersWithBigPasswordAgeDisabledPercentage%) are disabled."
  
      Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor White
      Write-Host $FindingText_EN -ForegroundColor Black -BackgroundColor White
  
      $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-text-users-with-big-password-age.txt" -Encoding utf8
      "`r`n" | Out-File -FilePath "$ReportingPath/finding-text-users-with-big-password-age.txt" -Append -Encoding utf8
      $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-text-users-with-big-password-age.txt" -Append -Encoding utf8

      $UsersWithBigPasswordAge | Out-File -FilePath "$ReportingPath/UsersWithBigPasswordAge.csv" -Encoding utf8
}

Write-Host '##############################################' -BackgroundColor Black
Write-Host '##               Admin Count                ##' -BackgroundColor Black
Write-Host '##############################################' -BackgroundColor Black
Write-Host 'Checking for users with AdminCount set...' -ForegroundColor Black -BackgroundColor White
# User Accounts with Admin Count
## Check if there are accounts that have an admin count set
$UsersWithAdminCount = $results.Users | Select-Object UserName, Enabled, AdminCount |
                                        Where-Object {$_.AdminCount -eq 1}

                                        
$UsersWithAdminCountCount = $($UsersWithAdminCount.Count)

if($UsersWithAdminCountCount -gt 0)
{
    $UsersWithAdminCountEnabled = $UsersWithAdminCount | Where-Object {$_.Enabled -eq $true}
    $UsersWithAdminCountDisabled = $UsersWithAdminCount | Where-Object {$_.Enabled -eq $false}
    
    $UsersWithAdminCountEnabledCount = $($UsersWithAdminCountEnabled.Count)
    $UsersWithAdminCountDisabledCount = $($UsersWithAdminCountDisabled.Count)

    $UsersWithAdminCountPercentage = $([math]::Round($UsersWithAdminCountCount / $TotalNumberOfUsers *100,2))
    $UsersWithAdminCountEnabledPercentage = $([math]::Round($UsersWithAdminCountEnabledCount / $UsersWithAdminCountCount *100,2))
    $UsersWithAdminCountDisabledPercentage = $([math]::Round($UsersWithAdminCountDisabledCount / $UsersWithAdminCountCount *100,2))
    
    Write-Host "Total users with Admin Count: $UsersWithAdminCountCount" -ForegroundColor Black -BackgroundColor DarkGray
    Write-Host "Total users with Admin Count (%): $UsersWithAdminCountPercentage% " -ForegroundColor Black -BackgroundColor DarkGray
    Write-Host "Users with Admin Count (Enabled): $UsersWithAdminCountEnabledCount " -ForegroundColor Black -BackgroundColor Red
    Write-Host "Users with Admin Count Enabled (%): $UsersWithAdminCountEnabledPercentage% " -ForegroundColor Black -BackgroundColor Red
    Write-Host "Users with Admin Count (Disabled): $UsersWithAdminCountDisabledCount" -ForegroundColor Black -BackgroundColor Yellow
    Write-Host "Users with Admin Count Disabled (%): $UsersWithAdminCountDisabledPercentage%" -ForegroundColor Black -BackgroundColor Yellow

    # No finding text, but an CSV export that we can use for the report
    $UsersWithAdminCount | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$ReportingPath/UsersWithAdminCount.csv"
}
else
{
    Write-Host "No users with AdminCount attribute set were identified"
}



Write-Host '#####################################################' -BackgroundColor Black
Write-Host '##           Cannot change password                ##' -BackgroundColor Black
Write-Host '#####################################################' -BackgroundColor Black
Write-Host 'Checking for users that cannot change their password' -ForegroundColor Black -BackgroundColor White
# User Accounts that cannot change passwords
## Check if there are accounts that cannot change passwords
$CannotChangePwdUsers = $results.Users | Select-Object UserName, Enabled, 'Cannot Change Password' | 
                                         Where-Object { $_.'Cannot Change Password' -eq $true}

$CannotChangePwdUsersCount = $(($CannotChangePwdUsers.Count))

if($CannotChangePwdUsersCount -gt 0)
{
    $CannotChangePwdUsersEnabled = $CannotChangePwdUsers | Where-Object { $_.'Enabled' -eq $true}
    $CannotChangePwdUsersDisabled = $CannotChangePwdUsers | Where-Object { $_.'Enabled' -eq $false}
    
    $CannotChangePwdUsersEnabledCount = $(($CannotChangePwdUsersEnabled.Count))
    $CannotChangePwdUsersDisabledCount = $(($CannotChangePwdUsersDisabled.Count))
    
    $CannotChangePwdUsersPercentage= $([math]::Round($CannotChangePwdUsersCount / $TotalNumberOfUsers * 100, 2))
    $CannotChangePwdUsersEnabledPercentage = $([math]::Round($CannotChangePwdUsersEnabledCount / $CannotChangePwdUsersCount * 100, 2))
    $CannotChangePwdUsersDisabledPercentage = $([math]::Round($CannotChangePwdUsersDisabledCount / $CannotChangePwdUsersCount * 100, 2))
    
    Write-Host "Total users that cannot change passwords: $CannotChangePwdUsersCount" -ForegroundColor Black -BackgroundColor DarkGray
    Write-Host "Total Users that cannot change passwords (%): $CannotChangePwdUsersPercentage%" -ForegroundColor Black -BackgroundColor DarkGray
    
    Write-Host "Total Users that cannot change passwords (Enabled): $CannotChangePwdUsersEnabledCount" -ForegroundColor Black -BackgroundColor Red
    Write-Host "Total Users that cannot change passwords (Enabled) (%): $CannotChangePwdUsersEnabledPercentage%" -ForegroundColor Black -BackgroundColor Red
    
    Write-Host "Total Users that cannot change passwords (Disabled): $CannotChangePwdUsersDisabledCount" -ForegroundColor Black -BackgroundColor Yellow
    Write-Host "Total Users that cannot change passwords (Disabled) (%): $CannotChangePwdUsersDisabledPercentage%"- -ForegroundColor Black -BackgroundColor Yellow

    $FindingText_DE="Während unserer Tests haben wir festgestellt, dass sich in der Domäne $domainName, welche insgesamt $([string]::Format('{0:N0}',$TotalNumberOfUsers)) Benutzerkonten verwaltet, $([string]::Format('{0:N0}',$CannotChangePwdUsersCount)) ($CannotChangePwdUsersPercentage%) Benutzerkonten keine Möglichkeit der Passwortänderung besitzen. Davon sind aktuell $([string]::Format('{0:N0}',$CannotChangePwdUsersEnabledCount)) ($CannotChangePwdUsersEnabledPercentage%) Benutzerkonten als aktiv gelistet und $([string]::Format('{0:N0}',$CannotChangePwdUsersDisabledCount)) ($CannotChangePwdUsersDisabledPercentage%) deaktiviert."
    $FindingText_EN="During our testing, we found that in the $domainName domain, which manages a total of $([string]::Format('{0:N0}',$TotalNumberOfUsers)) user accounts, $([string]::Format('{0:N0}',$CannotChangePwdUsersCount)) ($CannotChangePwdUsersPercentage%) user accounts do not have the ability to change their passwords. Of these, $([string]::Format('{0:N0}',$CannotChangePwdUsersEnabledCount)) ($CannotChangePwdUsersEnabledPercentage%) user accounts are currently listed as active and $CannotChangePwdUsersDisabledCount ($CannotChangePwdUsersDisabledPercentage%) are disabled."

    Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor White
    Write-Host $FindingText_EN -ForegroundColor Black -BackgroundColor White

    $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-text-cannot_chg_pwd.txt" -Encoding utf8
    "`r`n" | Out-File -FilePath "$ReportingPath/finding-text-cannot_chg_pwd.txt" -Append -Encoding utf8
    $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-text-cannot_chg_pwd.txt" -Append -Encoding utf8

    $CannotChangePwdUsers | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$ReportingPath/CannotChangePwdUsers.csv"
}
else
{
    Write-Host "No users were identified that have the cannot change their password attribute set" -ForegroundColor Black -BackgroundColor Green
}

Write-Host '##############################################' -BackgroundColor Black
Write-Host '##           PASSWD NOT REQD                ##' -BackgroundColor Black
Write-Host '##############################################' -BackgroundColor Black
Write-Host 'Checking for users that can have an empty password' -ForegroundColor Black -BackgroundColor White
# User Accounts that can have an empty password
## Check if there are accounts that can have an empty password (PASSWDNOTREQ)
$passwdNotReqUsers = $results.Users | Select-Object UserName, 
                                                    Enabled, 
                                                  'Password Not Required' |  
                                      Where-Object { $_.'Password Not Required' -eq $true}

$passwdNotReqUsersCount=$($passwdNotReqUsers.Name).count

if($passwdNotReqUsersCount -gt 0){
        
      $passwdNotReqUsersEnabled = $passwdNotReqUsers | Where-Object { $_.'Enabled' -eq $true}
      $passwdNotReqUsersDisabled = $passwdNotReqUsers | Where-Object { $_.'Enabled' -eq $false}

      $passwdNotReqUsersEnabledCount=$($passwdNotReqUsersEnabled.Name).Count
      $passwdNotReqUsersDisabledCount=$($passwdNotReqUsersDisabled.Name).Count 

      $passwdNotReqUsersPercentage = $([math]::Round(($passwdNotReqUsersCount / $TotalNumberOfUsers * 100),2))
      $passwdNotReqUsersEnabledPercentage = $([math]::Round(($passwdNotReqUsersEnabledCount / $passwdNotReqUsersCount * 100),2))
      $passwdNotReqUsersDisabledPercentage = $([math]::Round(($passwdNotReqUsersDisabledCount / $passwdNotReqUsersCount * 100),2))

      Write-Host "Total Users with Password Not Required attribute set: $passwdNotReqUsersCount" -BackgroundColor DarkGray -ForegroundColor Black
      Write-Host "Total Users with Password Not Required attribute set (%): $passwdNotReqUsersPercentage%" -BackgroundColor DarkGray -ForegroundColor Black

      Write-Host "Users with Password Not Required attribute set (Enabled): $passwdNotReqUsersEnabledCount" -BackgroundColor Red -ForegroundColor Black
      Write-Host "Users with Password Not Required attribute set (Enabled) (%): $passwdNotReqUsersEnabledPercentage%" -BackgroundColor Red -ForegroundColor Black

      Write-Host "Users with Password Not Required attribute set (Disabled): $passwdNotReqUsersDisabledCount" -BackgroundColor Yellow -ForegroundColor Black
      Write-Host "Users with Password Not Required attribute set (Disabled) (%): $passwdNotReqUsersDisabledPercentage%" -BackgroundColor Yellow -ForegroundColor Black

      $FindingText_DE="Während unserer Tests haben wir festgestellt, dass sich in der Domäne $domainName, welche insgesamt $([string]::Format('{0:N0}',$TotalNumberOfUsers)) Benutzerkonten verwaltet, $([string]::Format('{0:N0}',$passwdNotReqUsersCount)) ($passwdNotReqUsersPercentage%) Benutzerkonten das PASSWD_NOTREQD Flag gesetzt haben. Davon sind aktuell $([string]::Format('{0:N0}',$passwdNotReqUsersEnabledCount)) ($passwdNotReqUsersEnabledPercentage%) Benutzerkonten als aktiv gelistet und $([string]::Format('{0:N0}',$passwdNotReqUsersDisabledCount)) ($passwdNotReqUsersDisabledPercentage%) deaktiviert."
      $FindingText_EN="During our testing, we found that in the $domainName domain, which manages a total of $([string]::Format('{0:N0}',$TotalNumberOfUsers)) user accounts, $([string]::Format('{0:N0}',$passwdNotReqUsersCount)) ($passwdNotReqUsersPercentage%) user accounts have the PASSWD_NOTREQD flag set. Of these, $([string]::Format('{0:N0}',$passwdNotReqUsersEnabledCount)) ($passwdNotReqUsersEnabledPercentage%) user accounts are currently listed as active and $([string]::Format('{0:N0}',$passwdNotReqUsersDisabledCount)) ($passwdNotReqUsersDisabledPercentage%) are disabled."

      Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor White
      Write-Host $FindingText_EN -ForegroundColor Black -BackgroundColor White

      $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-text-passwd-not-reqd.txt" -Encoding utf8
      "`r`n" | Out-File -FilePath "$ReportingPath/finding-text-passwd-not-reqd.txt" -Append -Encoding utf8
      $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-text-cannot_change_pw.txt" -Append -Encoding utf8

      $passwdNotReqUsers | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$ReportingPath/PasswdNotReqUsers.csv"
      
} else {
      Write-Host "No users with PASSWD NOT REQD attribute set were identified" -ForegroundColor Black -BackgroundColor Green 
}

Write-Host '#####################################################' -BackgroundColor Black
Write-Host '##           Inactive Accounts                     ##' -BackgroundColor Black
Write-Host '#####################################################' -BackgroundColor Black
Write-Host 'Checking for users that did not log in for more than 90 days' -ForegroundColor Black -BackgroundColor White
# High number of inactive accounts
## Check if there are accounts that have have been inactive for a long time (90 days) [LastLogon > 90] or never logged in [LastLogon = 0]

$InactiveUsers = $results.Users | Select-Object UserName, 
                                                Enabled, 
                                                @{Name="LastLogon_days";Expression={[int] $_.'Logon Age (days)'}} | 
                                  Where-Object {($_.LastLogon_days -gt 90) -or ($_.LastLogon_days -eq 0)}
                                  
$InactiveUsersCount=$($InactiveUsers.Name).Count

If($InactiveUsersCount -gt 0)
{
    $InactiveUsersEnabled = $InactiveUsers | Where-Object {$_.Enabled -eq $true}
    $InactiveUsersDisabled = $InactiveUsers | Where-Object {$_.Enabled -eq $false}

    $InactiveUsersEnabledCount=$($InactiveUsersEnabled.Name).Count.ToString()
    $InactiveUsersDisabledCount=$($InactiveUsersDisabled.Name).Count.ToString()

    $InactiveUsersPercentage = $([math]::Round(($InactiveUsersCount / $TotalNumberOfUsers * 100),2))
    $InactiveUsersEnabledPercentage=$([math]::Round(($InactiveUsersEnabledCount / $InactiveUsersCount * 100),2))
    $InactiveUsersDisabledPercentage=$([math]::Round(($InactiveUsersDisabledCount / $InactiveUsersCount * 100),2))

    Write-Host "Total Users inactive for more than 90 days: $InactiveUsersCount" -BackgroundColor DarkGray -ForegroundColor Black
    Write-Host "Percentage of users that are inactive for more than 90 days: $InactiveUsersPercentage%" -BackgroundColor DarkGray -ForegroundColor Black

    Write-Host "Users inactive for more than 90 days (Enabled): $InactiveUsersEnabledCount" -BackgroundColor Red -ForegroundColor Black
    Write-Host "Percentage of users that are inactive for more than 90 days (Enabled): $InactiveUsersEnabledPercentage%" -BackgroundColor Red -ForegroundColor Black

    Write-Host "Users inactive for more than 90 days (Disabled): $InactiveUsersDisabledCount" -BackgroundColor Yellow -ForegroundColor Black
    Write-Host "Percentage of users that are inactive for more than 90 days (Disabled): $InactiveUsersDisabledPercentage%" -BackgroundColor Yellow -ForegroundColor Black

    $FindingText_DE="Wir haben festgestellt, dass sich in der Domäne $domainName von insgesamt $([string]::Format('{0:N0}',$TotalNumberOfUsers)) Benutzerkonten, $([string]::Format('{0:N0}',$InactiveUsersCount)) ($InactiveUsersPercentage%) Benutzerkonten seit mehr als 90 Tagen nicht mehr angemeldet haben. Davon sind aktuell $([string]::Format('{0:N0}',$InactiveUsersEnabledCount)) ($InactiveUsersEnabledPercentage%) als aktiv gelistet und $([string]::Format('{0:N0}',$InactiveUsersDisabledCount)) ($InactiveUsersDisabledPercentage%) sind deaktiviert."
    $FindingText_EN="We found that in the $domainName domain, out of a total of $([string]::Format('{0:N0}',$TotalNumberOfUsers)) user accounts, $([string]::Format('{0:N0}',$InactiveUsersCount)) ($InactiveUsersPercentage%) user accounts have not logged in for more than 90 days. Of these, $([string]::Format('{0:N0}',$InactiveUsersEnabledCount)) ($InactiveUsersEnabledPercentage%) are currently listed as active and $([string]::Format('{0:N0}',$InactiveUsersDisabledCount)) ($InactiveUsersDisabledPercentage%) are disabled."
    Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor White
    Write-Host $FindingText_EN -ForegroundColor Black -BackgroundColor White

    $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-text-inactive-users.txt" -Encoding utf8
    "`r`n" | Out-File -FilePath "$ReportingPath/finding-text-inactive-users.txt" -Append -Encoding utf8
    $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-text-inactive-users.txt" -Append -Encoding utf8

    $InactiveUsers | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$ReportingPath/UsersInactive.csv"
}
else
{
    Write-Host "No inactive users found" -ForegroundColor Black -BackgroundColor Green
}


Write-Host '#####################################################' -BackgroundColor Black
Write-Host '##           Inactive Computer                     ##' -BackgroundColor Black
Write-Host '#####################################################' -BackgroundColor Black
Write-Host 'Checking for computer that did not log in for more than 90 days' -ForegroundColor Black -BackgroundColor White
# Inactive computer accounts
## Check if there are computer accounts that have been inactive for a long time
$InactiveComputers = $results.Computers | Select-Object Name, 
                                                Enabled, 
                                                @{Name="LastLogon_days";Expression={[int] $_.'Logon Age (days)'}} | 
                                          Where-Object {$_.LastLogon_days -gt 90}

$InactiveComputersEnabled = $InactiveComputers | Where-Object {$_.Enabled -eq $true}
$InactiveComputersDisabled = $InactiveComputers | Where-Object {$_.Enabled -eq $false}

if($InactiveComputers.Count -gt 0)
{
      $InactiveComputersCount=$($InactiveComputers.Name).Count.ToString()
      $InactiveComputersPercentage=$([math]::Round(($InactiveComputersCount / $TotalNumberOfComputers * 100),2))
      $InactiveComputersEnabledCount=$($InactiveComputersEnabled.Name).Count.ToString()
      $InactiveComputersEnabledPercentage= $([math]::Round(($InactiveComputersEnabledCount / $InactiveComputersCount * 100),2))
      $InactiveComputersDisabledCount=$($InactiveComputersDisabled.Name).Count.ToString()
      $InactiveComputersDisabledPercentage= $([math]::Round(($InactiveComputersDisabledCount / $InactiveComputersCount * 100),2))

      Write-Host "Total Computer Accounts inactive for more than 90 days: $InactiveComputersCount" -BackgroundColor DarkGray -ForegroundColor Black
      Write-Host "Percentage of Computer Accounts inactive for more than 90 days: $InactiveComputersPercentage%" -BackgroundColor DarkGray -ForegroundColor Black
      Write-Host "Computer Accounts inactive for more than 90 days (Enabled): $InactiveComputersEnabledCount" -BackgroundColor Red -ForegroundColor Black
      Write-Host "Computer Accounts inactive for more than 90 days (Enabled) (%): $InactiveComputersEnabledPercentage%" -BackgroundColor Red -ForegroundColor Black
      Write-Host "Computer Accounts inactive for more than 90 days (Disabled): $InactiveComputersDisabledCount" -BackgroundColor Yellow -ForegroundColor Black
      Write-Host "Computer Accounts inactive for more than 90 days (Disabled) (%): $InactiveComputersDisabledPercentage%" -BackgroundColor Yellow -ForegroundColor Black

      $FindingText_DE="Während unserer Tests haben wir festgestellt, dass sich in der Domäne $domainName, welche insgesamt $([string]::Format('{0:N0}',$TotalNumberOfComputers)) Computersysteme verwaltet, $([string]::Format('{0:N0}',$InactiveComputersCount)) ($InactiveComputersPercentage%) Computer befinden, welche sich seit mehr wie 90 Tagen nicht mehr angemeldet haben. Davon sind aktuell $([string]::Format('{0:N0}',$InactiveComputersEnabledCount)) ($InactiveComputersEnabledPercentage%) aktiv und $([string]::Format('{0:N0}',$InactiveComputersDisabledCount)) ($InactiveComputersDisabledPercentage%) deaktiviert. Dies können entweder alte Einträge im AD Verzeichnisdienst sein, wofür die entsprechenden Systeme nicht mehr vorhanden sind oder anderweitige Ursachen haben."
      $FindingText_EN="During our tests, we found that in the $domainName domain, which manages a total of $([string]::Format('{0:N0}',$TotalNumberOfUsers)) computer systems, there are $([string]::Format('{0:N0}',$InactiveComputersCount)) ($InactiveComputersPercentage%) computers that have not logged in for more than 90 days. Of these, $([string]::Format('{0:N0}',$InactiveComputersEnabledCount)) ($InactiveComputersEnabledPercentage%) are currently active and $([string]::Format('{0:N0}',$InactiveComputersDisabledCount)) ($InactiveComputersDisabledPercentage%) are deactivated. These can either be old entries in the Active Directory, for which the corresponding systems no longer exist, or have other causes."
      Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor White
      Write-Host $FindingText_EN -ForegroundColor Black -BackgroundColor White
      $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-text-inactive-computers.txt" -Encoding utf8
      "`r`n" | Out-File -FilePath "$ReportingPath/finding-text-inactive-users-computers.txt" -Append -Encoding utf8
      $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-text-inactive-computers.txt" -Append -Encoding utf8

      $InactiveComputers | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$ReportingPath/InactiveComputers.csv"
}
else
{
    Write-Host "No computers inactive for more than 90 days" -ForegroundColor Black -BackgroundColor Green
}



# Unknown accounts
## Check if there are accounts that are unknown

# Risky SMB shares
## Check if there are shares that are risky - Use CME and parse with PS to deliver to the customer for evaluation

## Known AD Vulns
# Check if noPAC vuln exists

# Check if petitpotam vuln exists

Write-Host '#####################################################' -BackgroundColor Black
Write-Host '##           Accounts with DES Enabled             ##' -BackgroundColor Black
Write-Host '#####################################################' -BackgroundColor Black
Write-Host 'Checking for users that have DES enabled' -ForegroundColor Black -BackgroundColor White
# Check if accounts have DES enabled
$results.Users | where-object {$_.'Kerberos DES Only' -eq $true} | Format-Table
$DESUsers = $results.Users | where-object {$_.'Kerberos DES Only' -eq $true}


if($DESUsers.Count -gt 0)
{
      $DESUsersCount=$($DESUsers.Name).Count.ToString()
      $DESUsersEnabled = $results.Users | where-object {$_.'Kerberos DES Only' -eq $true} | where-object {$_.Enabled -eq $true}
      $DESUsersDisabled = $results.Users | where-object {$_.'Kerberos DES Only' -eq $true} | where-object {$_.Enabled -eq $false}
      $DESUsersEnabledCount=$($DESUsersEnabled.Name).Count.ToString()
      $DESUsersDisabledCount=$($DESUsersDisabled.Name).Count.ToString()
      $DESUsersPercentage=$([math]::Round(($DESUsersCount / $TotalNumberOfUsers * 100),2))
      $DESUsersEnabledPercentage=$([math]::Round(($DESUsersEnabledCount / $DESUsersCount * 100),2))
      $DESUsersDisabledPercentage=$([math]::Round(($DESUsersDisabledCount / $DESUsersCount * 100),2))

      Write-Host "Total Users with DES enabled: $DESUsersCount" -ForegroundColor Black -BackgroundColor DarkGray
      Write-Host "Percentage of Users with DES enabled: $DESUsersPercentage%" -ForegroundColor Black -BackgroundColor DarkGray
      Write-Host "Users with DES enabled (Enabled): $DESUsersEnabledCount" -ForegroundColor Black -BackgroundColor Red
      Write-Host "Users with DES enabled (Enabled) (%): $DESUsersEnabledPercentage%" -ForegroundColor Black -BackgroundColor Red
      Write-Host "Users with DES enabled (Disabled): $DESUsersDisabledCount" -ForegroundColor Black -BackgroundColor Yellow
      Write-Host "Users with DES enabled (Disabled) (%): $DESUsersDisabledPercentage%" -ForegroundColor Black -BackgroundColor Yellow

      $FindingText_DE="Während unserer Tests haben wir festgestellt, dass sich in der Domäne $domainName, welche insgesamt $([string]::Format('{0:N0}',$TotalNumberOfUsers)) Benutzer verwaltet, $([string]::Format('{0:N0}',$DESUsersCount)) ($DESUsersPercentage%) Benutzer befinden die DES aktiviert haben. Davon sind aktuell $([string]::Format('{0:N0}',$DESUsersEnabledCount)) ($DESUsersEnabledPercentage%) aktiv und $([string]::Format('{0:N0}',$DESUsersDisabledCount)) ($DESUsersDisabledPercentage%) deaktiviert. Dies können entweder alte Einträge im AD Verzeichnisdienst sein, wofür die entsprechenden Benutzer nicht mehr vorhanden sind oder anderweitige Ursachen haben."
      $FindingText_EN=""

      $FindingText_DE | Out-File -FilePath "$ReportingPath/finding-text-des-users.txt" -Encoding utf8
      "`r`n" | Out-File -FilePath "$ReportingPath/finding-text-des-users.txt" -Append -Encoding utf8
      $FindingText_EN | Out-File -FilePath "$ReportingPath/finding-text-des-users.txt" -Append -Encoding utf8
      
      $DESUsers | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$ReportingPath/DESUsers.csv"
      #Write-Host $FindingText_DE -ForegroundColor Black -BackgroundColor Magenta
}
else
{
    Write-Host "No users with DES enabled" -ForegroundColor Black -BackgroundColor Green
}

Write-Host '#########################################################' -BackgroundColor Black
Write-Host '##           Unsupported Operating Systems             ##' -BackgroundColor Black
Write-Host '#########################################################' -BackgroundColor Black
Write-Host 'Checking for computers that are enabled and have a unsupported operating system' -ForegroundColor Black -BackgroundColor White
# Check non-supported Windows Versions in the domain that are still active
$NonSupportedVersions = @('Windows NT',
                          'Windows XP',
                          'Windows 2000',
                          'Windows 2003',
                          'Windows Vista',
                          'Windows 2008',
                          'Windows 7',
                          'Windows 8',
                          'Windows Server 2000',
                          'Windows Server 2003',
                          'Windows Server 2008'
                          )

$OutdatedSystems= @()

foreach($version in $NonSupportedVersions) {
    $r=($results.Computers | Where-Object {($_.'Operating System' -match $version) -and ($_.Enabled -eq $true)})
    $OutdatedSystems += $r
}
$OutdatedSystemsCount=$OutdatedSystems.Count

if($OutdatedSystemsCount -gt 0)
{
    $OutdatedSystemsPercentage=$([math]::Round(($OutdatedSystemsCount / $TotalNumberOfComputers * 100),2))
    Write-Host "Total Computers with unsupported Operating Systems: $OutdatedSystemsCount" -ForegroundColor Black -BackgroundColor Red
    Write-Host "Percentage of Computers with unsupported Operating Systems: $OutdatedSystemsPercentage%" -ForegroundColor Black -BackgroundColor Red
    $OutdatedSystems | Format-Table -AutoSize

    $OutdatedSystems | Export-Csv -Path "$ReportingPath/OutdatedOSSystems.csv" -NoTypeInformation -Encoding UTF8
}
else
{
    Write-Host "No computers with unsupported Operating Systems" -ForegroundColor Black -BackgroundColor Green
}

Write-Host '#########################################################' -BackgroundColor Black
Write-Host '##           Use of Build-In Administrator              ##' -BackgroundColor Black
Write-Host '#########################################################' -BackgroundColor Black
Write-Host 'Checking for use of the Build-In Administrator account' -ForegroundColor Black -BackgroundColor White
# Check if native administrator account has been used recently
$BuildInAdmin = $results.Users | Select-Object UserName, 
                               Enabled, 
                               @{Name="Logon_age";Expression={[int]$_.'Logon Age (days)'}} | 
                 Where-Object {$_.UserName -eq 'Administrator'} |
                 Where-Object {$_.LogonAge_days -lt 30}

if($($BuildInAdmin.Name).Count -gt 0)
{
    Write-Host "The Build-In Administrator account has been used within the last 30 days" -ForegroundColor Black -BackgroundColor Red
    $BuildInAdmin
}
else
{
    Write-Host "The Build-In Administrator account has not been used in the last 30 days" -ForegroundColor Black -BackgroundColor Green
}
                

# Exchange Windows Permission - Change permissions on domain root (PingCastle)

Write-Host '#########################################################' -BackgroundColor Black
Write-Host '##           Admin accounts with old passwords         ##' -BackgroundColor Black
Write-Host '#########################################################' -BackgroundColor Black
Write-Host 'Checking for admins that have old passwords' -ForegroundColor Black -BackgroundColor White
# Admin accounts with old passwords
$AdminsWithOldPasswords=$results.Users | Select-Object UserName, 
                                                       Enabled, 
                                                       'Last Logon Date', 
                                                       'Password LastSet', 
                                                       @{Name="Password_age";Expression={[int]$_.'Password Age (days)'}},
                                                       AdminCount |
                                         Where-Object {$_.AdminCount -eq 1} |
                                         Where-Object {$_.Password_age -gt 365}
                                        
if($AdminsWithOldPasswords.Count -gt 0){
    $AdminsWithOldPasswordsCount=$AdminsWithOldPasswords.Count
    $AdminsWithOldPasswordsEnabledCount=$($AdminsWithOldPasswords | Where-Object {$_.Enabled -eq $true}).UserName.Count
    $AdminsWithOldPasswordsDisabledCount=$($AdminsWithOldPasswords | Where-Object {$_.Enabled -eq $false}).UserName.Count
    $AdminsWithOldPasswordsEnabledPercentage=$([math]::Round(($AdminsWithOldPasswordsEnabledCount / $AdminsWithOldPasswordsCount * 100),2))
    $AdminsWithOldPasswordsDisabledPercentage=$([math]::Round(($AdminsWithOldPasswordsDisabledCount / $AdminsWithOldPasswordsCount * 100),2))

    Write-Host "There are $($AdminsWithOldPasswords.Count) admin accounts with passwords older than 365 days" -ForegroundColor Black -BackgroundColor Red
    Write-Host "Admin accounts with old passwords (Enabled): $($AdminsWithOldPasswordsEnabledCount)" -ForegroundColor Black -BackgroundColor Red
    Write-Host "Admin accounts with old passwords (Enabled) (%): $($AdminsWithOldPasswordsEnabledPercentage)%" -ForegroundColor Black -BackgroundColor Red
    Write-Host "Admin accounts with old passwords (Disabled): $($AdminsWithOldPasswordsDisabledCount)" -ForegroundColor Black -BackgroundColor Yellow
    Write-Host "Admin accounts with old passwords (Disabled) (%): $($AdminsWithOldPasswordsDisabledPercentage)%" -ForegroundColor Black -BackgroundColor Yellow

    $AdminsWithOldPasswords | Format-Table -AutoSize
    $AdminsWithOldPasswords | Export-Csv -Path "$ReportingPath/AdminsWithOldPasswords.csv" -NoTypeInformation -Encoding UTF8
}
else
{
    Write-Host "No admin accounts with old passwords" -ForegroundColor Green -BackgroundColor Black
}

# Ensure that bogus Windows 2016 AD prep did not introduce vulnerabilities (PingCastle)

# Is the build in administrator account in use?
$results.Users | Where-Object {$_.UserName -eq "Administrator"} | 
                 Format-Table UserName, Enabled, 'Last Logon Date', 'Password LastSet'

# Show Members of the privileged groups (Famous 13)
$DNSAdminsSID=($results.Groups | Where-Object {$_.Name -eq "DNSAdmins"}).SID

$AdminGroups = @{'Domain Admins' = "$domainSID-512"
                 'Enterprise Admins' = "$domainSID-519"
                 'Administrators' = "S-1-5-32-544"
                 'Schema Admins'= "$domainSID-518"
                 'Print Operators'= "S-1-5-32-544"
                 'Server Operators'= "S-1-5-32-549"
                 'Account Operators'= "S-1-5-32-548"
                 'Backup Operators'= "S-1-5-32-551"
                 'Certificate Administrators'= "$domainSID-517"
                 'Group Policy Creator Owners'= "$domainSID-520"
                 'Remote Desktop Users'= "S-1-5-32-555"
                 'DnsAdmins'= "$DNSAdminsSID"
                }


# Active enum using toolz (Connection required)
# Loading toolset
iex(New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/PowerShellMafia/PowerSploit/dev/Recon/PowerView.ps1")
iex(New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/S3cur3Th1sSh1t/PowerSharpPack/master/PowerSharpBinaries/Invoke-LdapSignCheck.ps1")
iex(New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/NotMedic/NetNTLMtoSilverTicket/master/Get-SpoolStatus.ps1")
iex(New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/kfosaaen/Get-LAPSPasswords/master/Get-LAPSPasswords.ps1")
iex(New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/michiiii/Check-SMBSigning/master/Check-SMBSigning.ps1")
iex(New-Object Net.WebClient).DownloadString("https://raw.githubusercontent.com/michiiii/SmbScanner/master/Check-SMBv1.ps1")



Write-Host '#########################################################' -BackgroundColor Black
Write-Host '##           Unusual Primary Group IDs                 ##' -BackgroundColor Black
Write-Host '#########################################################' -BackgroundColor Black
Write-Host 'Checking for unusual primary group ids' -ForegroundColor Black -BackgroundColor White
$UsersWithUnusalPrimaryGroupID = Get-Netuser | Where-Object {$_.primarygroupid -ne 513} | Where-Object {($_.samaccountname -ne "Administrator") -and ($_.samaccountname -ne "Gast") -and ($_.samaccountname -ne "Guest")} | Format-Table samaccountname, primarygroupID

if($UsersWithUnusalPrimaryGroupID.Count -gt 0){
    Write-Host "There are $($UsersWithUnusalPrimaryGroupID.Count) users with unusual primary group IDs" -ForegroundColor Black -BackgroundColor Red
    $UsersWithUnusalPrimaryGroupID | Format-Table -AutoSize
} else
{
    Write-Host "No users with unusual primary group IDs" -ForegroundColor Black -BackgroundColor Green
}

Write-Host '#########################################################' -BackgroundColor Black
Write-Host '##           Domain Controller Print Spooler           ##' -BackgroundColor Black
Write-Host '#########################################################' -BackgroundColor Black
Write-Host 'Checking Domain Controllers that expose print spoolers' -ForegroundColor Black -BackgroundColor White


ForEach ($DC in (Get-NetDomainController)) {
  Get-SpoolStatus $DC.Name
}

Write-Host '#########################################################' -BackgroundColor Black
Write-Host '##               Checking LDAP Signing                 ##' -BackgroundColor Black
Write-Host '#########################################################' -BackgroundColor Black
Write-Host 'Checking for LDAP Signing requirements' -ForegroundColor Black -BackgroundColor White
$domainUserName= Read-Host -Prompt "Enter your user name (without Domain suffix): "
$domainUserPassword= Read-Host -Prompt "Enter the password for the user: "
Invoke-LdapSignCheck -Command "-domain $($domainName) -user $($domainUserName) -password $($domainUserPassword)" -ErrorAction SilentlyContinue

Write-Host '#########################################################' -BackgroundColor Black
Write-Host '##           SeMachineAccountPrivilege                ##' -BackgroundColor Black
Write-Host '#########################################################' -BackgroundColor Black
Write-Host 'Checking for groups that have the SeMachineAccounntPrivilege privilege' -ForegroundColor Black -BackgroundColor White
$SeMachineAccountPrivilege=(Get-DomainPolicy -Policy DC).PrivilegeRights.SEMachineAccountPrivilege

ForEach($sid in $SeMachineAccountPrivilege)
{
  $resource=(Convert-SidToName $sid.Substring(1))
  if($resource -eq "Authenticated Users")
  {
    Write-Host "$($resource) has the SeMachineAccountPrivileges privileges" -ForegroundColor Black -BackgroundColor Red
  }
  else
  {
    Write-Host "$($resource) has the SeMachineAccountPrivileges privileges" -ForegroundColor Black -BackgroundColor Yellow
  }
}


# Static creds for local computers
## Check LAPS 
Write-Host '##################################' -BackgroundColor Black
Write-Host '##              LAPS            ##' -BackgroundColor Black
Write-Host '##################################' -BackgroundColor Black
Write-Host 'Checking for LAPS implementation' -ForegroundColor Black -BackgroundColor White

$LAPSComputersCount = (Get-LAPSPasswords | Where-Object {$_.Store -eq 1}).count
$NonLAPSComputersCount = ((Get-Netcomputer).Count)-$LAPSComputersCount

$ComputerCountEnabled= $(($results.Computers | Where-Object {$_.Enabled -eq $true}).count).ToString()
$ComputerCountDisabled= $(($results.Computers | Where-Object {$_.Enabled -eq $false}).count).ToString()

Write-Host "Total number of computers in Active Directory: $TotalNumberOfComputers"
Write-Host "Number of enabled computer accounts: $ComputerCountEnabled"
Write-Host "Number of disabled computer accounts: $ComputerCountDisabled"
Write-Host "Number of computers with LAPS: $LAPSComputersCount"
Write-Host "Number of computers without LAPS: $NonLAPSComputersCount"

$LAPSPercentage = [math]::Round(($LAPSComputersCount / $ComputerCountEnabled * 100),2)

if($LAPSPercentage -ge 90)
{
  Write-Host "LAPS is implemented on $LAPSPercentage % of the computers in the $domainName domain" -ForegroundColor Black -BackgroundColor Green
}
elseif (($LAPSPercentage -ge 65) -and ($LAPSPercentage -lt 90))
{
  Write-Host "LAPS is implemented on $LAPSPercentage % of the computers in the $domainName domain" -ForegroundColor Black -BackgroundColor Yellow
}
elseif ($LAPSPercentage -lt 65){
  Write-Host "LAPS is implemented on $LAPSPercentage % of the computers in the $domainName domain" -ForegroundColor Black -BackgroundColor Red
}

Get-LAPSPasswords | Where-Object {$_.Store -eq 1}

Write-Host '#########################################' -BackgroundColor Black
Write-Host '##              SMB-Signing            ##' -BackgroundColor Black
Write-Host '#########################################' -BackgroundColor Black
$activeComputers=Get-NetComputer | Select-Object -Property Name,
                                                           DNSHostName,
                                                           @{Name="LogonAge";Expression={[int](New-TimeSpan -Start ($_.lastlogontimestamp) -End (Get-Date)).Days}} |
                                   Where-Object {($_.LogonAge -lt 30) -and 
                                                 ($_.DNSHostName -ne $null)}

ForEach($computer in $activeComputers){
    if((Test-Port -ComputerName $computer.DNSHostName -Port 445).PortOpened){
      Check-SMBSigning -Target $computer.DNSHostName
    }else{
      Write-Host "$($computer.DNSHostName) is not reachable" -ForegroundColor Black -BackgroundColor Gray
    }   
}

Write-Host '#########################################' -BackgroundColor Black
Write-Host '##              SMBv1            ##' -BackgroundColor Black
Write-Host '#########################################' -BackgroundColor Black
$activeComputers=Get-NetComputer | Select-Object -Property Name,
                                                           DNSHostName,
                                                           @{Name="LogonAge";Expression={[int](New-TimeSpan -Start ($_.lastlogontimestamp) -End (Get-Date)).Days}} |
                                   Where-Object {($_.LogonAge -lt 30) -and 
                                                 ($_.DNSHostName -ne $null)}

ForEach($computer in $activeComputers){
    if((Test-Port -ComputerName $computer.DNSHostName -Port 445).PortOpened){
      if(([PingCastle.Scanners.SmbScanner]::CheckSMBv1($computer.DNSHostName)) -like "*enabled*"){
          Write-Host "$($computer.DNSHostName) has SMBv1 enabled" -ForegroundColor Black -BackgroundColor Red
      }else
      {   
          Write-Host "$($computer.DNSHostName) has SMBv1 disabled" -ForegroundColor Black -BackgroundColor Green
      }

    }else{
      Write-Host "$($computer.DNSHostName) is not reachable" -ForegroundColor Black -BackgroundColor Gray
    }   
}

# Last AD Database Backup

# Retrival of LAPS Password
## https://azurecloudai.blog/2019/10/01/laps-security-concern-computers-joiners-are-able-to-see-laps-password/
## https://www.securityinsider-wavestone.com/2020/01/taking-over-windows-workstations-pxe-laps.html
## [MITRE]T1555.005 Credentials from Password Stores: Password Managers

# LAPS details
## Get the LAPS details from the domaim

# DNS
## Unauthenticated users can create DNS records
## Authenticated user can create DNS records

# WSUS

# WEF

# Login Scripts
# What are the login scripts
# Can unpriv users run them or write them

# NetCease



#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion
#Script Execution goes here
#Log-Finish -LogPath $sLogFile

# Write a function that will return a custom PS object array with all findings. The function takes the ADRecon results as input.

# Write a function that calculated the difference between two dates. The function takes two dates as input. The function returns the difference in days.
function Get-DifferenceBetweenDates($date1, $date2) {
  return (New-TimeSpan -Start ($date1) -End ($date2)).Days
}
