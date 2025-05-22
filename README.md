# PowerChell

A proof-of-concept aimed at creating **a PowerShell console in C/C++**, with all the security features patched or disabled: Antimalware Scan Interface (AMSI), Script Block logging, Module logging, Transcription, Execution Policy, and Constrained Language Mode (CLM).

## Build

1. Open the solution file `PowerChell.sln` with Visual Studio (you must have the Windows SDK installed).
2. In the toolbar, select `RELEASE-EXE` if you want to build the executable (.exe) file, or `RELEASE-DLL` if you want to build the DLL. In both cases, the target configuration will be `x64` because this is the only supported platform.
3. In the top bar, click `Build > Build Solution` to build the project.

## Usage

### Open a PowerShell Console

You should be able to run the **executable** straight away:

```console
C:\Users\Dummy\Downloads>PowerChell.exe
Windows PowerChell
Copyright (C) Microsoft Corporation. All rights reserved.

PS C:\Users\Dummy\Downloads> $PSVersionTable

Name                           Value
----                           -----
PSVersion                      5.1.26100.2161
PSEdition                      Desktop
PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0...}
BuildVersion                   10.0.26100.2161
CLRVersion                     4.0.30319.42000
WSManStackVersion              3.0
PSRemotingProtocolVersion      2.3
SerializationVersion           1.1.0.1
```

As for the **DLL**, you can use the following command:

```batch
rundll32 PowerChell.dll,Start
```

In the command above, `Start` is the name of a dummy function. It exists only to prevent `rundll32` from complaining about not finding the entry point. You can very well specify any entry point you want. It will work as long as you don't close the error dialog.

### Execute a Command

You can also execute a PowerShell command like this:

```console
C:\Users\Dummy\Downloads>PowerChell.exe -c "$PSVersionTable"

+-----------------------------------+
| POWERSHELL STANDARD OUTPUT STREAM |
+-----------------------------------+

Name                           Value
----                           -----
PSVersion                      5.1.26100.3624
PSEdition                      Desktop
PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0...}
BuildVersion                   10.0.26100.3624
CLRVersion                     4.0.30319.42000
WSManStackVersion              3.0
PSRemotingProtocolVersion      2.3
SerializationVersion           1.1.0.1
```

This also works with the **DLL**, albeit less convient because you won't see the result:

```batch
rundll32 PowerChell.dll,Start -c "$PSVersionTable"
```

## Syntax Highlighting and other Goodies

In PowerShell v3+, Microsoft added a built-in module named [`PSReadLine`](https://github.com/PowerShell/PSReadLine) to bring a bash-like experience to the PowerShell console. PowerChell doesn't load any module by design, so if you want to enable features such as syntax highlighting and other goodies that this module implements, you need to load it manually.

However, this has a **side effect in regard to detection by EDR**. In doing so, you will also enable PowerShell command logging to the current uesr's console history file `%APPDATA%\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt`. This file is actively monitored by some security solutions (see [issue #1](https://github.com/scrt/PowerChell/issues/1)).

To load this module, and disable console history, you should do the following.

```powershell
Import-Module PSReadLine
Set-PSReadLineOption -HistorySaveStyle SaveNothing
```

![Image](https://github.com/user-attachments/assets/5d7979a9-9b52-4403-ab8d-88d13962bf73)

## Caveats

- If you open any of the source files and Visual Studio is screaming at you because it can't find the `mscorlib` stuff, that's expected. You need to build the solution at least once. It will generate the `mscorlib.tlh` file automatically.
- The code of the DLL will likely need to be adapted if you want it to work properly using DLL sideloading.

## Authors

- Clément Labro
    - Mastodon: [https://infosec.exchange/@itm4n](https://infosec.exchange/@itm4n)
    - GitHub: [https://github.com/itm4n](https://github.com/itm4n)

## Credit

There would be many resources, blog posts, and tools to credit. Unfortunately, I haven't kept track of all them, but here are the main ones.

**Tools**

- https://github.com/calebstewart/bypass-clm
- https://github.com/anonymous300502/Nuke-AMSI
- https://github.com/OmerYa/Invisi-Shell
- https://github.com/leechristensen/UnmanagedPowerShell
- https://github.com/racoten/BetterNetLoader
- https://gist.github.com/Arno0x/386ebfebd78ee4f0cbbbb2a7c4405f74 (`loadDotNetAssemblyFromMemory.cpp`)
- https://gist.github.com/tandasat/e595c77c52e13aaee60e1e8b65d2ba32 (`KillETW.ps1`)

**Blog posts**

- [Unmanaged .NET Patching](https://www.outflank.nl/blog/2024/02/01/unmanaged-dotnet-patching/)
- [Massaging your CLR: Preventing Environment.Exit in In-Process .NET Assemblies](https://www.mdsec.co.uk/2020/08/massaging-your-clr-preventing-environment-exit-in-in-process-net-assemblies/)
- [15 Ways to Bypass the PowerShell Execution Policy](https://www.netspi.com/blog/technical-blog/network-pentesting/15-ways-to-bypass-the-powershell-execution-policy/)
