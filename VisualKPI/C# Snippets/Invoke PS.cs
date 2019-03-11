            var pss = @"C:\Users\IGNITE\Desktop\csharp.ps1";//path to the ps script
            Process cmd = new Process();//create a CMD to run powershell, directly running powershell didn't seem to work
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.Start();
            cmd.StandardInput.WriteLine("Powershell Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;");//This enables execution of powershell scripts 
            cmd.StandardInput.WriteLine("PowerShell.exe -windowstyle hidden " + pss + " ");//Runs a powershell script without window
            cmd.StandardInput.Flush();
            cmd.StandardInput.Close();
            cmd.WaitForExit();
