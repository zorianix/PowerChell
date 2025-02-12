#pragma once
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <stdio.h>
#include <metahost.h>
#include <propvarutil.h>
#pragma comment(lib, "mscoree.lib")
#pragma comment(lib, "Propsys.lib")

#import <mscorlib.tlb> raw_interfaces_only \
    high_property_prefixes("_get","_put","_putref") \
    rename("ReportEvent", "InteropServices_ReportEvent") \
    rename("or", "InteropServices_or")
using namespace mscorlib;

#define APP_DOMAIN L"PowerChell"
#define ASSEMBLY_NAME_SYSTEM_CORE L"System.Core"
#define ASSEMBLY_NAME_SYSTEM_RUNTIME L"System.Runtime"
#define ASSEMBLY_NAME_MICROSOFT_POWERSHELL_CONSOLEHOST L"Microsoft.PowerShell.ConsoleHost"
#define ASSEMBLY_NAME_SYSTEM_MANAGEMENT_AUTOMATION L"System.Management.Automation"

typedef struct _CLR_CONTEXT
{
    ICLRMetaHost* pMetaHost;
    ICLRRuntimeInfo* pRuntimeInfo;
    ICorRuntimeHost* pRuntimeHost;
    IUnknown* pAppDomainThunk;
} CLR_CONTEXT, * PCLR_CONTEXT;

BOOL InitializeCommonLanguageRuntime(PCLR_CONTEXT pClrContext, mscorlib::_AppDomain** ppAppDomain);
void DestroyCommonLanguageRuntime(PCLR_CONTEXT pClrContext, mscorlib::_AppDomain* pAppDomain);
BOOL FindAssemblyPath(LPCWSTR pwszAssemblyName, LPWSTR* ppwszAssemblyPath);
BOOL GetAssembly(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, mscorlib::_Assembly** ppAssembly);
BOOL LoadAssembly(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, mscorlib::_Assembly** ppAssembly);
BOOL FindMethodInArray(SAFEARRAY* pMethods, LPCWSTR pwszMethodName, LONG lNbArgs, mscorlib::_MethodInfo** ppMethodInfo);
BOOL PrepareMethod(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtMethodHandle);
BOOL GetFunctionPointer(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtMethodHandle, PULONG_PTR pFunctionPointer);
BOOL GetJustInTimeMethodAddress(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, LPCWSTR pwszClassName, LPCWSTR pwszMethodName, DWORD dwNbArgs, PULONG_PTR pMethodAddress);