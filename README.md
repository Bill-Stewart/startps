# startps

## AUTHOR

Bill Stewart - bstewart at iname dot com

## LICENSE

**startps** is covered by the GNU Public License (GPL). See the file `LICENSE` for details.

## DOWNLOAD

https://github.com/Bill-Stewart/startps/releases/

## SYNOPSIS

Runs a PowerShell script or opens an interactive PowerShell session.

## USAGE

**startps** [_parameters_] [_scriptfile_ [**--** _script params_]]  
or: **startps** [_parameters_] **--interactive**

Parameter names are case-sensitive. Long parameter names can be specified partially if enough of the parameter name is specified to prevent ambiguity.

## GENERAL PARAMETERS

The following parameters apply whether running a script or not:

Long Name                        | Short Name      | Description
-------------------------------- | --------------- | -------------------------------------------
**--configurationname=**_name_   |                 | Run PS in a configuration endpoint
**--consolefilename=**_filename_ |                 | Load a PS console file
**--core**[**=**_path_]          | **-c** [_path_] | Use PS Core (_path_ = path to `pwsh.exe`)
**--disableexecutionpolicy**     | **-D**          | Disable PS execution policy
**--elevate**                    | **-e**          | Request to run as administrator
**--help**                       | **-h**          | Display usage information
**--modulepath=**_path_          | **-m** _path_   | Appends _path_ to **PSModulePath**
**--mta**                        |                 | Use multi-threaded apartment
**--outputformat=**_format_      |                 | Specify output format (**Text** or **XML**)
**--quiet**                      | **-q**          | Suppress error messages
**--sta**                        |                 | Use single-threaded apartment
**--version=**_version_          |                 | Run PS using specified version
**--windowstyle=**_style_        | **-W** _style_  | Specify a window style
**--windowtitle**[**=**_text_]   | **-t** [_text_] | Specify a window title
**--workingdirectory=**_path_    | **-d** _path_   | Specify a working directory

Notes:

* If a parameter's argument contains spaces, enclose it in `"` characters. For example:

  * `--core="C:\Program Files\PowerShell\7\pwsh.exe"`
  * `--windowtitle="Sample window title"`
  * `--workingdirectory="C:\Program Files"`

* If a parameter's argument contains the `"` character, enclose the argument in `"` characters and double the embedded `"` characters. For example:

  * `--windowtitle="Sample ""quoted"" string"`

* The **--outputputformat** parameter's argument must be **Text** or **XML** (the argument is not case-sensitive).

* The **--windowstyle** (**-W**) parameter's argument must be one of the following: **Normal**, **Minimized**, **Maximized**, **Hidden**, **NormalNotActive**, or **MinimizedNotActive** (the argument is not case-sensitive),

* Windows PowerShell ignores the **--workingdirectory** (**-d**) parameter if **--elevate** (**-e**) is specified.

* The **--mta** and **--sta** parameters are mutually exclusive.

* The **--consolefilename** and **--version** parameters are specific to Windows PowerShell and are ignored if running PowerShell Core.

* In practice, the **--version** parameter is only used to start the Windows PowerShell 2.0 engine (i.e., `--version 2`; not recommended).

## PARAMETERS IF RUNNING A SCRIPT

The following parameters apply only when running a script:

Long Name              | Short Name | Description
---------------------- | ---------- | -------------------------------------------
**--loadprofile**      |            | Load PS profile(s) before running script
**--noexit**           |            | Keep PS window open after running script
**--noninteractive**   | **-n**     | Run the script non-interactively
**--pause**            | **-p**     | Pause window after script completes
**--scriptinexedir**   | **-s**     | Script is in same directory as executable
**--wait**             | **-w**     | Wait for exit and return process exit code
_scriptfile_           |            | Path/filename of script to run
**--** _script params_ |            | Put script parameters (if any) after **--**

Notes:

* If the script file's path or filename contains spaces, enclose it in `"` characters; e.g.: `"C:\My Scripts\script.ps1"`.

* **--pause** (**-p**) and **--wait** (**-w**) are ignored if **--noexit** is specified.

* You can use the **--scriptinexedir** (**-s**) parameter if you place the script file in the same directory as `startps.exe`. This is useful in scenarios where you don't know the full path to the script file (i.e., in a startup or logon script).

## PARAMETERS IF RUNNING INTERACTIVELY

The following parameters apply only when running PS interactively:

Long Name         | Short Name | Description
----------------- | ---------- | ------------------------------------
**--interactive** | **-i**     | Open interactive PS window
**--logo**        |            | Show the copyright banner at startup
**--noprofile**   |            | Do not load PS profile(s)

## EXIT CODES

* Exit code will be 0 for success, non-zero for error

* If **--wait** (**-w**) specified, exit code will be PowerShell process exit code

## POWERSHELL CORE

By default, **startps** executes Windows PowerShell rather than PowerShell Core. You can run PowerShell Core instead by specifying the **--core** (**-c**) parameter. If you have more than one version of PowerShell Core installed, you can run a specific version by specifying the path and filename of `pwsh.exe` as the argument; e.g.: `--core="C:\Program Files\PowerShell\7\pwsh.exe"`.

Alternatively, you can set the `PSPath` environment variable to the path of the directory where `pwsh.exe` is installed. If the `PSPath` variable is set, **startps** will use the `pwsh.exe` executable in that directory. If the `PSPath` environment variable is defined, it takes precedence over all other attempts at finding `powershell.exe` or `pwsh.exe`.

## 32-BIT VS. 64-BIT

It's recommended to run the 64-bit (x86_64) version of **startps** on 64-bit operating systems unless there's a specific need on that platform to run a 32-bit (i386) version of PowerShell. Running the 32-bit (i386) version of **startps** on a 64-bit operating system will, of course, run the 32-bit version of Windows PowerShell or PowerShell Core (if the 32-bit version of PowerShell Core is installed).

## EXAMPLES

* Run a PowerShell script using Windows PowerShell and pause window after script completes:

      startps --pause "C:\Script Files\Test Script.ps1"

  `--pause` can be abbreviated as `-p`.

* Start an interactive PowerShell Core session:

      startps --core --interactive

* Run a PowerShell script using PowerShell Core, passing parameters to the script:

      startps -c "C:\Script Files\Core Script.ps1" -- -Param1 "Script Param"

  The `--` parameter specifies that everything after it is parameters for the script.
