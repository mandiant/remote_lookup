
Contributed by FireEye FLARE Team
Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
License: GPL

Article link: 
  https://www.fireeye.com/blog/threat-research/2017/06/remote-symbol-resolution.html

This is a small tool which can scan a 32bit process and build an
export name/address map which can be queried.

Precompiled binaries can be found in the /bin folder.

It supports single searches, bulk lookups from file, or requests 
from network clients.

Sample remote clients are provided in Python, C#, VB6 and D.

The tool supports the following input formats:
    hexMemoryAddress,
    case insensitive api name
    ws2_32@13,
    ntdll!atoi or msvcrt.atoi

This application has the following dependencies:
  - sppe.dll     - PE File Format Library 
  - procLib.dll  - Process Library
  - MSWINSCK.OCX - Microsoft Winsock ActiveX control.

If run as administrator the application can register these 
itself on the first run. The machine will also require the 
VB6 runtimes which are already pre installed on most systems.

The source for the other libraries can be found here:
  https://github.com/dzzie/libs/tree/master/pe_lib
  https://github.com/dzzie/libs/tree/master/proc_lib

Note: 
-------------------------------------------------------------
proclib does support 64bit processes and addresses however 
64bit support has not yet been added to this tool.



