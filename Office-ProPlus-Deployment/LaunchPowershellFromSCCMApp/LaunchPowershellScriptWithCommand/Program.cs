using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace LaunchPowershellScriptWithCommand
{
    class Program
    {
        static void Main(string[] args)
        {
            //couple of things to note about this program
            //1. The powershell script you wish to run must not require an arguments, and must not require a function call
            //2. To launch, the call will look like the following: C:\temp>LaunchPowershellScriptWithCommand.exe "& { . .\Change-OfficeChannel.ps1; Change-OfficeChannel -Channel Deferred }"
            //
            Process p = new Process();
            p.StartInfo.FileName = "Powershell.exe";
            p.StartInfo.Arguments = @"-ExecutionPolicy Bypass -NoExit -Command " + args[0];
            p.Start();
            p.Close();
        }
    }
}
