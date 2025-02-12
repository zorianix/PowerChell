#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include "../PowerChellLib/powershell.h"

#ifndef _WINDLL
int main()
{
    StartPowerShell();

    return 0;
}
#else
extern "C" __declspec(dllexport) void APIENTRY Start();
void StartPowerShellInNewConsole();

HANDLE g_hEvent = NULL;

#pragma warning(disable : 4100)
BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        DisableThreadLibraryCalls(hModule);
        CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)StartPowerShellInNewConsole, NULL, 0, NULL);
        break;
    case DLL_PROCESS_DETACH:
        break;
    }

    return TRUE;
}

void APIENTRY Start()
{
    WaitForSingleObject(g_hEvent, INFINITE);
}

void StartPowerShellInNewConsole()
{
    if ((g_hEvent = CreateEventW(NULL, TRUE, FALSE, NULL)) != NULL)
    {
        AllocConsole();
        StartPowerShell();
        SetEvent(g_hEvent);
        FreeConsole();
    }
}
#endif