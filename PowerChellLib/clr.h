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
#define ASSEMBLY_NAME_SYSTEM_REFLECTION L"System.Reflection"
#define ASSEMBLY_NAME_SYSTEM_RUNTIME L"System.Runtime"
#define ASSEMBLY_NAME_MICROSOFT_POWERSHELL_CONSOLEHOST L"Microsoft.PowerShell.ConsoleHost"
#define ASSEMBLY_NAME_SYSTEM_MANAGEMENT_AUTOMATION L"System.Management.Automation"
#define BINDING_FLAGS_PUBLIC_STATIC mscorlib::BindingFlags::BindingFlags_Public | mscorlib::BindingFlags::BindingFlags_Static
#define BINDING_FLAGS_PUBLIC_INSTANCE mscorlib::BindingFlags::BindingFlags_Public | mscorlib::BindingFlags::BindingFlags_Instance
#define BINDING_FLAGS_NONPUBLIC_STATIC mscorlib::BindingFlags::BindingFlags_NonPublic | mscorlib::BindingFlags::BindingFlags_Static
#define BINDING_FLAGS_NONPUBLIC_INSTANCE mscorlib::BindingFlags::BindingFlags_NonPublic | mscorlib::BindingFlags::BindingFlags_Instance

typedef struct _CLR_CONTEXT
{
    ICLRMetaHost* pMetaHost;
    ICLRRuntimeInfo* pRuntimeInfo;
    ICorRuntimeHost* pRuntimeHost;
    IUnknown* pAppDomainThunk;
} CLR_CONTEXT, * PCLR_CONTEXT;

namespace clr
{
    BOOL InitializeCommonLanguageRuntime(PCLR_CONTEXT pClrContext, mscorlib::_AppDomain** ppAppDomain);
    void DestroyCommonLanguageRuntime(PCLR_CONTEXT pClrContext, mscorlib::_AppDomain* pAppDomain);
    BOOL FindAssemblyPath(LPCWSTR pwszAssemblyName, LPWSTR* ppwszAssemblyPath);
    BOOL GetAssembly(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, mscorlib::_Assembly** ppAssembly);
    BOOL LoadAssembly(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, mscorlib::_Assembly** ppAssembly);
    BOOL CreateInstance(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, LPCWSTR pwszClassName, VARIANT* pvtInstance);
    BOOL GetType(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, LPCWSTR pwszTypeFullName, mscorlib::_Type** ppType);
    BOOL GetProperty(mscorlib::_Type* pType, mscorlib::BindingFlags bindingFlags, LPCWSTR pwszPropertyName, mscorlib::_PropertyInfo** ppPropertyInfo);
    BOOL GetPropertyValue(mscorlib::_Type* pType, mscorlib::BindingFlags bindingFlags, VARIANT vtObject, LPCWSTR pwszPropertyName, VARIANT* pvtPropertyValue);
    BOOL GetField(mscorlib::_Type* pType, mscorlib::BindingFlags bindingFlags, LPCWSTR pwszFieldName, mscorlib::_FieldInfo** ppFieldInfo);
    BOOL GetFieldValue(mscorlib::_Type* pType, mscorlib::BindingFlags bindingFlags, VARIANT vtObject, LPCWSTR pwszFieldName, VARIANT* pvtFieldValue);
    BOOL GetMethod(mscorlib::_Type* pType, mscorlib::BindingFlags bindingFlags, LPCWSTR pwszMethodName, LONG lNbArg, mscorlib::_MethodInfo** ppMethodInfo);
    BOOL InvokeMethod(mscorlib::_MethodInfo* pMethodInfo, VARIANT vtObject, SAFEARRAY* pParameters, VARIANT* pvtResult);
    BOOL FindMethodInArray(SAFEARRAY* pMethods, LPCWSTR pwszMethodName, LONG lNbArgs, mscorlib::_MethodInfo** ppMethodInfo);
    BOOL PrepareMethod(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtMethodHandle);
    BOOL GetFunctionPointer(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtMethodHandle, PULONG_PTR pFunctionPointer);
    BOOL GetJustInTimeMethodAddress(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, LPCWSTR pwszClassName, LPCWSTR pwszMethodName, DWORD dwNbArgs, PULONG_PTR pMethodAddress);
}

namespace dotnet
{
    BOOL System_Object_GetType(mscorlib::_AppDomain* pAppDomain, VARIANT vtObject, VARIANT* pvtObjectType);
    BOOL System_Type_GetProperty(mscorlib::_AppDomain* pAppDomain, VARIANT vtTypeObject, LPCWSTR pwszPropertyName, VARIANT* pvtPropertyInfo);
    BOOL System_Reflection_PropertyInfo_GetValue(mscorlib::_AppDomain* pAppDomain, VARIANT vtPropertyInfo, VARIANT vtObject, SAFEARRAY* pIndex, VARIANT* pvtValue);
    BOOL System_Reflection_PropertyInfo_GetValue(mscorlib::_AppDomain* pAppDomain, VARIANT vtPropertyInfo, VARIANT vtObject, VARIANT* pvtValue);
}