# PowerShell Runspace Functions (That Actually Work in 2025)

## Why This Exists

In PowerShell, 7.x `ForEach-Object -Parallel` exists which is nice for simple, limited scripting. However, it becomes apparently painful when trying to import your own custom functions, sharing variables, or even loading installed PS modules.

I had some old runspace scripts from way back that worked great in PowerShell 5.1, but when I tried running them in PowerShell 7.x? No joy. Half the methods changed, some stuff got deprecated. 
Looked around GitHub for existing runspace solutions and found a bunch from like 2018-2020 that were either abandoned or didn't work properly with modern PowerShell. So I built my own. 🤷‍♂️

## What's Different

This actually lets you:
- Import your custom functions into runspaces (not just built-in cmdlets)
- Share variables between your main script and the parallel jobs
- Import modules properly without weird errors
- See pretty progress bars instead of just... waiting and hoping
- Handle timeouts without everything exploding
- Get meaningful error messages when stuff breaks

To use it, you will need some knowledge in runspace concepts. If you do, load the functions (just dot source it):
. .\Runspace-Functions.ps1

Refer to the example script to help with how to apply it to your own circumstances.
