# COMProxy

A COM client and server for testing COM hijack proxying. If you are running a COM hijack, proxying the legitimate COM server may result in better stability, thats the idea around this PoC. 

This also provides an example in-process COM server DLL that can be taken and modified for your tooling. This implementation will read whatever COM CLSID is being hijacked, attempt to find the legitimate path in `HKLM\Software\Classes\CLSID`, and proxy the COM interface so the COM client receives the expected pointers.

### How to setup the test case?

Build both projects. The TestCOMClient will create a `WScript.Shell` object and run calc. The TestCOMServer is a COM server DLL that will start a thread printing to the screen every second.

Run this in a .reg file for the hijack:

```
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Classes\CLSID\{72C24DD5-D70A-438B-8A42-98424B88AFB8}]
@="Test Hijack"

[HKEY_CURRENT_USER\Software\Classes\CLSID\{72C24DD5-D70A-438B-8A42-98424B88AFB8}\InprocServer32]
@="C:\\Users\\user\\Desktop\\TestCOMServer.dll"
"ThreadingModel"="Apartment"
```
