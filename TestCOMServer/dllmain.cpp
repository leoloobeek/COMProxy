#include <Windows.h>
#include <comutil.h> // #include for _bstr_t
#include <string>

BOOL FindOriginalCOMServer(wchar_t* GUID, wchar_t* DLLName);
DWORD MyThread();
typedef HRESULT(__stdcall *_DllGetClassObject)(REFCLSID rclsid, REFIID riid, LPVOID* ppv);

UINT g_uThreadFinished;
extern UINT g_uThreadFinished;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
	{
		AllocConsole();
		FILE *stream;
		freopen_s(&stream, "CONOUT$", "w+", stdout);
		g_uThreadFinished = 0;
		printf("[*] DLL_PROCESS_ATTACH\n");
		break;
	}
	case DLL_PROCESS_DETACH:
		printf("[*] DLL_PROCESS_DETACH\n");
		break;
	}
	return TRUE;
}

STDAPI DllCanUnloadNow(void)
{
	wprintf(L"[+] DllCanUnloadNow\n");
	// Ensure our thread can finish before we're unloaded
	do 
	{
		Sleep(1);
	} while (g_uThreadFinished == 0);
	
	wprintf(L"[+] All done, exiting.\n");
	return S_OK;
}
STDAPI DllRegisterServer(void)
{
	wprintf(L"[+] DllRegisterServer\n");
	return S_OK;
}
STDAPI DllUnregisterServer(void)
{
	wprintf(L"[+] DllUnregisterServer\n");
	return S_OK;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	wprintf(L"[+] DllGetClassObject\n");
	CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)MyThread, NULL, 0, NULL);
	HMODULE hDLL;
	_DllGetClassObject lpGetClassObject;

	LPOLESTR lplpsz;
	HRESULT hResult = StringFromCLSID(rclsid, &lplpsz);
	wchar_t* DLLName = new wchar_t[MAX_PATH];

	if (!FindOriginalCOMServer((wchar_t*)lplpsz, DLLName))
	{
		wprintf(L"[-] Couldn't find original COM server\n");
		return S_FALSE;
	}

	wprintf(L"[+] Found original COM server: %s\n", DLLName);

	// Load up the original COM server
	hDLL = LoadLibrary(DLLName);
	if (hDLL == NULL)
	{
		wprintf(L"[-] hDLL was NULL\n");
		return S_FALSE;
	}

	// Find the DllGetClassObject for original COM server
	lpGetClassObject = (_DllGetClassObject)GetProcAddress(hDLL, "DllGetClassObject");
	if (lpGetClassObject == NULL)
	{
		wprintf(L"[-] lpGetClassObject is null\n");
		return S_FALSE;
	}

	// Call the intended DllGetClassObject from original COM server
	// This will get all the necessary pointers and should be all set if successful
	HRESULT hr = lpGetClassObject(rclsid, riid, ppv);
	if FAILED(hr) 
	{
		wprintf(L"[-] lpGetClassObject got hr 0x%08lx\n", hr);
	}

	wprintf(L"[+] Done!\n");
	
	return S_OK;
}

BOOL FindOriginalCOMServer(wchar_t* GUID, wchar_t* DLLName)
{
	HKEY hKey;
	HKEY hCLSIDKey;
	DWORD nameLength = MAX_PATH;

	wprintf(L"[*] Beginning search for GUID %s\n", GUID);
	LONG lResult = RegOpenKeyExW(HKEY_LOCAL_MACHINE, (LPCWSTR)L"SOFTWARE\\Classes\\CLSID", 0, KEY_READ, &hKey);
	if (lResult != ERROR_SUCCESS) {
		wprintf(L"[-] Error getting CLSID path\n");
		return FALSE;
	}

	// Make sure HKLM\Software\Classes\CLSID\{GUID} exists
	lResult = RegOpenKeyExW(hKey, GUID, 0, KEY_READ, &hCLSIDKey);
	if (lResult != ERROR_SUCCESS) {
		wprintf(L"[-] Error getting GUID path\n");
		RegCloseKey(hKey);
		return FALSE;
	}

	// Read the value of HKLM's InProcServer32
	lResult = RegGetValueW(hCLSIDKey, (LPCWSTR)L"InProcServer32", NULL, RRF_RT_ANY, NULL, (PVOID)DLLName, &nameLength);
	if (lResult != ERROR_SUCCESS) {
		wprintf(L"[-] Error getting InProcServer32 value: %d\n", lResult);
		RegCloseKey(hKey);
		RegCloseKey(hCLSIDKey);
		return FALSE;
	}

	return TRUE;
}

DWORD MyThread()
{
	printf("[+] MyThread\n\n");

	for (int i = 0; i < 30; i++)
	{
		printf("[*] %d\n", i);
		Sleep(1000);
	}

	g_uThreadFinished = 1;
	return 0;
}

