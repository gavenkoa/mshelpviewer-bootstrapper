
MS HelpViewer Bootstrapper
==========================

# Intent

MS Help Viewer 2.x creates a storage for a given MSVC under some so called "Catalogs":

```
Get-ChildItem -Path "HKLM:SOFTWARE\WOW6432Node\Microsoft\Help" |
  % { Get-ChildItem -path $_.PSPath } |
  ? PSChildName -eq "Catalogs" |
  % { Get-ChildItem -path $_.PSPath }


Hive: HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Help\v2.2\Catalogs

Name            Property
----            --------
VisualStudio14  LocationPath : %ProgramData%\Microsoft\HelpLibrary2\Catalogs\VisualStudio14\
```

This helper scripts adds necessary files and registry keys to create an independent catalog with
custom `LocationPath` (several GiB storage for docs).

# Usage

* Run `mshelpviewer-bootstrapper.ps1` elevated (as Admin).
* Choose HelpViewer version (among installed).
* Choose catalog name.
* Select document root folder (catalog name will be appended but you free to edit path).
* Push "Create" button.

As a result you'll get a shortcut (`.lnk` file) on the Desktop to launch HelpViewer with your
catalog selected.

Launch shortcut and navigate to "Manage Content" tab. "Online" installation source is fixed &
predefined, to alter source select "Disk" and:

* select a local dir
* paste a URL and hit Enter to refresh list of available documents:

    * MSVC 2017/2019/2022: https://services.mtps.microsoft.com/ServiceAPI/catalogs/dev15/en-us
    * MSVC 2015: https://services.mtps.microsoft.com/ServiceAPI/catalogs/dev14/en-us
    * MSVC 2013: https://services.mtps.microsoft.com/ServiceAPI/catalogs/VisualStudio12/en-us
    * MSVC 2012: https://services.mtps.microsoft.com/ServiceAPI/catalogs/VisualStudio11/en-us
    * SQL Server 2016/2017/2019: https://services.mtps.microsoft.com/ServiceAPI/catalogs/sql2016/en-us

You might find another public URLs, I found above here:

    https://services.mtps.microsoft.com/ServiceAPI/catalogs

Few other links, could be sniffed delving into:

    https://services.mtps.microsoft.com/ServiceAPI/products

like:

https://services.mtps.microsoft.com/serviceapi/products/Dd776353/Dd776354/books/Dd253661/en-us
:   .NET Framework 3.5.

https://services.mtps.microsoft.com/serviceapi/products/Dd776353/Dd776354/books/Dd776355/en-us
:   .NET Framework 4

https://services.mtps.microsoft.com/serviceapi/products/Dd776353/Gg593679/books/Gg594376/en-us
:   Windows Driver Kit (2012).

# Supplementary

https://stackoverflow.com/questions/4701193/download-windows-api-reference-msdn-for-offline-use
:   Links for old CHM & offline MSDN `.iso` downloads.

https://github.com/nickdalt/VSHD/
:   Download documentation from `https://services.mtps.microsoft.com/ServiceAPI/catalogs` in a form,
    acceptable by HelpViewer "Manage Content" => "Disk" for offline installation!

https://marketplace.visualstudio.com/items?itemName=Moataz99.MSDNtoUSB
:   Backs up locally downloaded MSDN in HelpViewer format, later restore on other PC. Comes with
    Microsoft Help Viewer 2.3 installer inside - no need to download 10GiB of Visual Studio.

