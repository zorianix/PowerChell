#include "common.h"
#include "clr.h"

BOOL InitializeCommonLanguageRuntime(PCLR_CONTEXT pClrContext, mscorlib::_AppDomain** ppAppDomain)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    ICLRMetaHost* pMetaHost = NULL;
    ICLRRuntimeInfo* pRuntimeInfo = NULL;
    ICorRuntimeHost* pRuntimeHost = NULL;
    IUnknown* pAppDomainThunk = NULL;
    BOOL bIsLoadable;

    mscorlib::_AppDomain* pAppDomain = NULL;

    hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_ICLRMetaHost, reinterpret_cast<PVOID*>(&pMetaHost));
    EXIT_ON_HRESULT_ERROR(L"CLRCreateInstance", hr);

    hr = pMetaHost->GetRuntime(L"v4.0.30319", IID_ICLRRuntimeInfo, reinterpret_cast<PVOID*>(&pRuntimeInfo));
    EXIT_ON_HRESULT_ERROR(L"IMetaHost->GetRuntime", hr);

    hr = pRuntimeInfo->IsLoadable(&bIsLoadable);
    EXIT_ON_HRESULT_ERROR(L"IRuntimeInfo->IsLoadable", hr);

    hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_ICorRuntimeHost, reinterpret_cast<PVOID*>(&pRuntimeHost));
    EXIT_ON_HRESULT_ERROR(L"IRuntimeInfo->GetInterface", hr);

    hr = pRuntimeHost->Start();
    EXIT_ON_HRESULT_ERROR(L"IRuntimeHost->Start", hr);

    hr = pRuntimeHost->CreateDomain(APP_DOMAIN, nullptr, &pAppDomainThunk);
    EXIT_ON_HRESULT_ERROR(L"IRuntimeHost->CreateDomain", hr);

    hr = pAppDomainThunk->QueryInterface(IID_PPV_ARGS(&pAppDomain));
    EXIT_ON_HRESULT_ERROR(L"IAppDomainThunk->QueryInterface", hr);

    pClrContext->pMetaHost = pMetaHost;
    pClrContext->pRuntimeInfo = pRuntimeInfo;
    pClrContext->pRuntimeHost = pRuntimeHost;
    pClrContext->pAppDomainThunk = pAppDomainThunk;
    *ppAppDomain = pAppDomain;
    bResult = TRUE;

exit:
    if (!bResult && pAppDomain) pAppDomain->Release();
    if (!bResult && pAppDomainThunk) pAppDomainThunk->Release();
    if (!bResult && pRuntimeHost) pRuntimeHost->Release();
    if (!bResult && pRuntimeInfo) pRuntimeInfo->Release();
    if (!bResult && pMetaHost) pMetaHost->Release();

    return bResult;
}

void DestroyCommonLanguageRuntime(PCLR_CONTEXT pClrContext, mscorlib::_AppDomain* pAppDomain)
{
    if (pAppDomain) pAppDomain->Release();
    if (pClrContext->pAppDomainThunk) pClrContext->pAppDomainThunk->Release();
    if (pClrContext->pRuntimeHost) pClrContext->pRuntimeHost->Release();
    if (pClrContext->pRuntimeInfo) pClrContext->pRuntimeInfo->Release();
    if (pClrContext->pMetaHost) pClrContext->pMetaHost->Release();
}

BOOL FindAssemblyPath(LPCWSTR pwszAssemblyName, LPWSTR* ppwszAssemblyPath)
{
    BOOL bResult = FALSE;
    WIN32_FIND_DATA ffd = { 0 };
    LPCWSTR pwszAssemblyFolderPath = L"C:\\Windows\\Microsoft.NET\\assembly\\GAC_MSIL";
    LPWSTR pwszSearchPath = NULL;
    LPWSTR pwszAssemblyPath = NULL;
    HANDLE hFind = NULL;
    HANDLE hAssemblyFile = NULL;

    pwszSearchPath = (LPWSTR)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, MAX_PATH * sizeof(*pwszSearchPath));
    EXIT_ON_WIN32_ERROR(L"HeapAlloc", pwszSearchPath == NULL);

    pwszAssemblyPath = (LPWSTR)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, MAX_PATH * sizeof(*pwszAssemblyPath));
    EXIT_ON_WIN32_ERROR(L"HeapAlloc", pwszAssemblyPath == NULL);

    swprintf_s(pwszSearchPath, MAX_PATH, L"%ws\\%ws\\*", pwszAssemblyFolderPath, pwszAssemblyName);
    
    hFind = FindFirstFileW(pwszSearchPath, &ffd);
    EXIT_ON_WIN32_ERROR(L"FindFirstFileW", hFind == INVALID_HANDLE_VALUE);

    do
    {
        if (ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            swprintf_s(pwszAssemblyPath, MAX_PATH, L"%ws\\%ws\\%ws\\%ws.dll", pwszAssemblyFolderPath, pwszAssemblyName, ffd.cFileName, pwszAssemblyName);
            hAssemblyFile = CreateFileW(pwszAssemblyPath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
            if (hAssemblyFile != INVALID_HANDLE_VALUE)
            {
                bResult = TRUE;
                CloseHandle(hAssemblyFile);
                break;
            }
        }
    } while (FindNextFileW(hFind, &ffd) != 0);

    if (!bResult) goto exit;

    *ppwszAssemblyPath = pwszAssemblyPath;

exit:
    if (hFind && hFind != INVALID_HANDLE_VALUE) FindClose(hFind);
    if (!bResult && pwszAssemblyPath) HeapFree(GetProcessHeap(), HEAP_ZERO_MEMORY, pwszAssemblyPath);
    if (pwszSearchPath) HeapFree(GetProcessHeap(), HEAP_ZERO_MEMORY, pwszSearchPath);

    return bResult;
}

BOOL GetAssembly(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, mscorlib::_Assembly** ppAssembly)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    LONG lAssembliesLowerBound, lAssembliesUpperBound;
    LONG lTargetAssemblyNameLength, lCurrentAssemblyNameLength, lAssemblyNameIndex;
    BSTR bstrAssemblyFullName;
    SAFEARRAY* pLoadedAssembliesArray = NULL;
    mscorlib::_Assembly** ppLoadedAssemblies = NULL;

    lTargetAssemblyNameLength = (LONG)wcslen(pwszAssemblyName);

    hr = pAppDomain->GetAssemblies(&pLoadedAssembliesArray);
    EXIT_ON_HRESULT_ERROR(L"AppDomain->GetAssemblies", hr);

    SafeArrayGetLBound(pLoadedAssembliesArray, 1, &lAssembliesLowerBound);
    SafeArrayGetUBound(pLoadedAssembliesArray, 1, &lAssembliesUpperBound);

    hr = SafeArrayAccessData(pLoadedAssembliesArray, (void**)&ppLoadedAssemblies);
    EXIT_ON_HRESULT_ERROR(L"SafeArrayAccessData", hr);

    for (int i = 0; i < lAssembliesUpperBound - lAssembliesLowerBound + 1; i++)
    {
        bstrAssemblyFullName = NULL;
        hr = ppLoadedAssemblies[i]->get_FullName(&bstrAssemblyFullName);

        if (SUCCEEDED(hr))
        {
            //
            // If the target assembly name is 'Some.Assembly.Name', check whether
            // the current assembly's full name starts with 'Some.Assembly.Name,'.
            // As an example, the full name of the 'System.Management.Automation'
            // assembly would be something like:
            //   System.Management.Automation, Version=3.0.0.0, Culture=neutral, 
            //   PublicKeyToken=31bf3856ad364e35
            //

            lCurrentAssemblyNameLength = (LONG)wcslen(bstrAssemblyFullName);

            if (lCurrentAssemblyNameLength > lTargetAssemblyNameLength)
            {
                for (lAssemblyNameIndex = 0; lAssemblyNameIndex < lTargetAssemblyNameLength; lAssemblyNameIndex++)
                {
                    if (pwszAssemblyName[lAssemblyNameIndex] != bstrAssemblyFullName[lAssemblyNameIndex])
                        break;
                }

                if (lAssemblyNameIndex == lTargetAssemblyNameLength)
                {
                    if (bstrAssemblyFullName[lAssemblyNameIndex] == L',')
                    {
                        bResult = TRUE;
                        *ppAssembly = ppLoadedAssemblies[i];
                        break;
                    }
                }
            }

            SysFreeString(bstrAssemblyFullName);
        }
    }

exit:
    if (pLoadedAssembliesArray) SafeArrayDestroy(pLoadedAssembliesArray);

    return bResult;
}

BOOL LoadAssembly(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, mscorlib::_Assembly** ppAssembly)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    DWORD dwBytesRead;
    HANDLE hAssemblyFile = NULL;
    LPWSTR pwszAssemblyPath = NULL;
    LARGE_INTEGER liAssemblyFileSize = { 0 };
    SAFEARRAYBOUND sab = { 0 };
    SAFEARRAY* pSafeAssembly = NULL;
    mscorlib::_Assembly* pAssembly = NULL;

    // Check if assembly is already loaded
    if (!GetAssembly(pAppDomain, pwszAssemblyName, &pAssembly))
    {
        if (!FindAssemblyPath(pwszAssemblyName, &pwszAssemblyPath))
            goto exit;

        hAssemblyFile = CreateFileW(pwszAssemblyPath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        EXIT_ON_WIN32_ERROR(L"CreateFileW", hAssemblyFile == INVALID_HANDLE_VALUE);
        EXIT_ON_WIN32_ERROR(L"GetFileSizeEx", GetFileSizeEx(hAssemblyFile, &liAssemblyFileSize) == FALSE);

        sab = { (ULONG)liAssemblyFileSize.QuadPart, 0 };
        pSafeAssembly = SafeArrayCreate(VT_UI1, 1, &sab);

        EXIT_ON_WIN32_ERROR(L"ReadFile", ReadFile(hAssemblyFile, pSafeAssembly->pvData, (DWORD)liAssemblyFileSize.QuadPart, &dwBytesRead, NULL) == FALSE);

        hr = pAppDomain->Load_3(pSafeAssembly, &pAssembly);
        EXIT_ON_HRESULT_ERROR(L"AppDomain->Load_3", hr);
    }

    *ppAssembly = pAssembly;
    bResult = TRUE;

exit:
    if (pSafeAssembly) SafeArrayDestroy(pSafeAssembly);
    if (hAssemblyFile && hAssemblyFile != INVALID_HANDLE_VALUE) CloseHandle(hAssemblyFile);
    if (pwszAssemblyPath) HeapFree(GetProcessHeap(), HEAP_ZERO_MEMORY, pwszAssemblyPath);

    return bResult;
}

BOOL FindMethodInArray(SAFEARRAY* pMethods, LPCWSTR pwszMethodName, LONG lNbArgs, mscorlib::_MethodInfo** ppMethodInfo)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    LONG lMethodsLowerBound, lMethodsUpperBound;
    LONG lParametersLowerBound, lParametersUpperBound;
    BSTR bstrMethodName;
    SAFEARRAY* pParameters = NULL;
    mscorlib::_MethodInfo** ppMethods = NULL;
    mscorlib::_MethodInfo* pTargetMethodInfo = NULL;

    SafeArrayGetLBound(pMethods, 1, &lMethodsLowerBound);
    SafeArrayGetUBound(pMethods, 1, &lMethodsUpperBound);

    hr = SafeArrayAccessData(pMethods, (void**)&ppMethods);
    EXIT_ON_HRESULT_ERROR(L"SafeArrayAccessData", hr);

    for (int i = 0; i < lMethodsUpperBound - lMethodsLowerBound + 1; i++)
    {
        bstrMethodName = NULL;
        hr = ppMethods[i]->get_name(&bstrMethodName);

        if (SUCCEEDED(hr))
        {
            if (_wcsicmp(bstrMethodName, pwszMethodName) == 0)
            {
                hr = ppMethods[i]->GetParameters(&pParameters);

                if (SUCCEEDED(hr))
                {
                    SafeArrayGetLBound(pParameters, 1, &lParametersLowerBound);
                    SafeArrayGetUBound(pParameters, 1, &lParametersUpperBound);

                    if (lParametersUpperBound - lParametersLowerBound + 1 == lNbArgs)
                    {
                        pTargetMethodInfo = ppMethods[i];
                        break;
                    }
                }
            }

            SysFreeString(bstrMethodName);
        }
    }

    if (!pTargetMethodInfo)
    {
        PRINT_ERROR("Could not find a method named '%ws' with %d arguments.\n", pwszMethodName, lNbArgs);
        goto exit;
    }

    bResult = TRUE;
    *ppMethodInfo = pTargetMethodInfo;

exit:
    if (pParameters) SafeArrayDestroy(pParameters);

    return bResult;
}

BOOL PrepareMethod(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtMethodHandle)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    LONG lArgumentIndex;
    BSTR bstrRuntimeHelpersFullName = SysAllocString(L"System.Runtime.CompilerServices.RuntimeHelpers");
    SAFEARRAY* pRuntimeHelpersMethods = NULL;
    SAFEARRAY* pPrepareMethodArguments = NULL;
    VARIANT vtEmpty = { 0 };
    VARIANT vtResult = { 0 };
    mscorlib::_Assembly* pRuntimeAssembly = NULL;
    mscorlib::_Type* pRuntimeHelpersType = NULL;
    mscorlib::_MethodInfo* pPrepareMethod = NULL;

    if (!LoadAssembly(pAppDomain, ASSEMBLY_NAME_SYSTEM_RUNTIME, &pRuntimeAssembly))
        goto exit;

    hr = pRuntimeAssembly->GetType_2(bstrRuntimeHelpersFullName, &pRuntimeHelpersType);

    EXIT_ON_HRESULT_ERROR(L"(System.Runtime.CompilerServices.RuntimeHelpers) Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"System.Runtime.CompilerServices.RuntimeHelpers type", pRuntimeHelpersType);

    hr = pRuntimeHelpersType->GetMethods(
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Static | mscorlib::BindingFlags::BindingFlags_Public),
        &pRuntimeHelpersMethods
    );

    EXIT_ON_HRESULT_ERROR(L"(System.Runtime.CompilerServices.RuntimeHelpers) Type->GetMethods", hr);

    FindMethodInArray(pRuntimeHelpersMethods, L"PrepareMethod", 1, &pPrepareMethod);

    if (!pPrepareMethod)
    {
        PRINT_ERROR("Method RuntimeHelpers.PrepareMethod(RuntimeMethodHandle) not found.\n");
        goto exit;
    }

    pPrepareMethodArguments = SafeArrayCreateVector(VT_VARIANT, 0, 1);

    lArgumentIndex = 0;
    hr = SafeArrayPutElement(pPrepareMethodArguments, &lArgumentIndex, pvtMethodHandle);
    
    // The first parameter is supposed to be an object instance, but here we invoke
    // a static method, so it doesn't matter, and can just be "empty".
    hr = pPrepareMethod->Invoke_3(vtEmpty, pPrepareMethodArguments, &vtResult);
    EXIT_ON_HRESULT_ERROR(L"(RuntimeHelpers.GetFunctionPointer) MethodInfo->Invoke_3", hr);

    bResult = TRUE;

exit:
    if (bstrRuntimeHelpersFullName) SysFreeString(bstrRuntimeHelpersFullName);

    if (pRuntimeHelpersMethods) SafeArrayDestroy(pRuntimeHelpersMethods);
    if (pPrepareMethodArguments) SafeArrayDestroy(pPrepareMethodArguments);

    if (pPrepareMethod) pPrepareMethod->Release();
    if (pRuntimeHelpersType) pRuntimeHelpersType->Release();
    if (pRuntimeAssembly) pRuntimeAssembly->Release();

    return bResult;
}

BOOL GetFunctionPointer(mscorlib::_AppDomain* pAppDomain, VARIANT* pvtMethodHandle, PULONG_PTR pFunctionPointer)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    BSTR bstrRuntimeMethodHandleFullName = SysAllocString(L"System.RuntimeMethodHandle");
    SAFEARRAY* pRuntimeMethodHandleMethods = NULL;
    VARIANT vtFunctionPointer = { 0 };
    mscorlib::_Assembly* pRuntimeAssembly = NULL;
    mscorlib::_Type* pRuntimeMethodHandleType = NULL;
    mscorlib::_MethodInfo* pGetFunctionPointerInfo = NULL;

    if (!LoadAssembly(pAppDomain, L"System.Runtime", &pRuntimeAssembly))
        goto exit;

    hr = pRuntimeAssembly->GetType_2(bstrRuntimeMethodHandleFullName, &pRuntimeMethodHandleType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"RuntimeMethodHandle type", pRuntimeMethodHandleType);

    hr = pRuntimeMethodHandleType->GetMethods(
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Public | mscorlib::BindingFlags::BindingFlags_Instance),
        &pRuntimeMethodHandleMethods
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetMethods", hr);

    if (!FindMethodInArray(pRuntimeMethodHandleMethods, L"GetFunctionPointer", 0, &pGetFunctionPointerInfo))
        goto exit;

    hr = pGetFunctionPointerInfo->Invoke_3(*pvtMethodHandle, NULL, &vtFunctionPointer);
    EXIT_ON_HRESULT_ERROR(L"MethodInfo->Invoke_3", hr);

    *pFunctionPointer = vtFunctionPointer.ullVal;
    bResult = TRUE;

exit:
    if (bstrRuntimeMethodHandleFullName) SysFreeString(bstrRuntimeMethodHandleFullName);

    if (pRuntimeMethodHandleMethods) SafeArrayDestroy(pRuntimeMethodHandleMethods);

    if (pGetFunctionPointerInfo) pGetFunctionPointerInfo->Release();
    if (pRuntimeMethodHandleType) pRuntimeMethodHandleType->Release();
    if (pRuntimeAssembly) pRuntimeAssembly->Release();

    return bResult;
}

BOOL GetJustInTimeMethodAddress(mscorlib::_AppDomain* pAppDomain, LPCWSTR pwszAssemblyName, LPCWSTR pwszClassName, LPCWSTR pwszMethodName, DWORD dwNbArgs, PULONG_PTR pMethodAddress)
{
    BOOL bResult = FALSE;
    HRESULT hr;
    BSTR bstrClassName = SysAllocString(pwszClassName);
    BSTR bstrMethodInfoFullName = SysAllocString(L"System.Reflection.MethodInfo");
    BSTR bstrMethodHandlePropName = SysAllocString(L"MethodHandle");
    SAFEARRAY* pMethods = NULL;
    VARIANT vtMethodHandlePtr = { 0 };
    VARIANT vtMethodHandleVal = { 0 };
    mscorlib::_Assembly* pAssembly = NULL;
    mscorlib::_Assembly* pReflectionAssembly = NULL;
    mscorlib::_Type* pType = NULL;
    mscorlib::_Type* pMethodInfoType = NULL;
    mscorlib::_MethodInfo* pTargetMethodInfo = NULL;
    mscorlib::_PropertyInfo* pMethodHandlePropInfo = NULL;

    // Here, we include as many binding flags as we can so that we can list
    // ALL the methods of the target class.
    mscorlib::BindingFlags flags = static_cast<mscorlib::BindingFlags>(
        mscorlib::BindingFlags::BindingFlags_Instance |
        mscorlib::BindingFlags::BindingFlags_Static |
        mscorlib::BindingFlags::BindingFlags_Public |
        mscorlib::BindingFlags::BindingFlags_NonPublic |
        mscorlib::BindingFlags::BindingFlags_DeclaredOnly
    );

    if (!LoadAssembly(pAppDomain, pwszAssemblyName, &pAssembly))
        goto exit;

    if (!LoadAssembly(pAppDomain, L"System.Reflection", &pReflectionAssembly))
        goto exit;

    hr = pAssembly->GetType_2(bstrClassName, &pType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"Input type", pType);

    hr = pType->GetMethods(flags, &pMethods);
    EXIT_ON_HRESULT_ERROR(L"Type->GetMethods", hr);

    if (!FindMethodInArray(pMethods, pwszMethodName, dwNbArgs, &pTargetMethodInfo))
        goto exit;

    if (!pTargetMethodInfo)
    {
        PRINT_ERROR("Method '%ws.%ws' not found.\n", pwszClassName, pwszMethodName);
        goto exit;
    }

    //
    // The method for obtaining the MethodHandle from the MethodInfo object is
    // taken from this article.
    // 
    // Credit:
    //   - https://www.outflank.nl/blog/2024/02/01/unmanaged-dotnet-patching/
    //

    hr = pReflectionAssembly->GetType_2(bstrMethodInfoFullName, &pMethodInfoType);
    EXIT_ON_HRESULT_ERROR(L"Assembly->GetType_2", hr);
    EXIT_ON_NULL_POINTER(L"MethodInfo type", pMethodInfoType);

    hr = pMethodInfoType->GetProperty(
        bstrMethodHandlePropName,
        static_cast<mscorlib::BindingFlags>(mscorlib::BindingFlags::BindingFlags_Instance | mscorlib::BindingFlags::BindingFlags_Public),
        &pMethodHandlePropInfo
    );

    EXIT_ON_HRESULT_ERROR(L"Type->GetProperty", hr);
    EXIT_ON_NULL_POINTER(L"MethodHandle property type", pMethodHandlePropInfo);

    vtMethodHandlePtr.vt = VT_UNKNOWN;
    vtMethodHandlePtr.punkVal = pTargetMethodInfo;

    hr = pMethodHandlePropInfo->GetValue(vtMethodHandlePtr, NULL, &vtMethodHandleVal);
    EXIT_ON_HRESULT_ERROR(L"MethodHandle->GetValue", hr);

    //
    // Next, we can invoke 'RuntimeHelpers.PrepareMethod' to make sure the target
    // method is JIT-compiled, and finally get its effective address.
    //
    // Credit:
    //   - https://github.com/calebstewart/bypass-clm
    //   - https://www.mdsec.co.uk/2020/08/massaging-your-clr-preventing-environment-exit-in-in-process-net-assemblies/
    //

    if (!PrepareMethod(pAppDomain, &vtMethodHandleVal))
        goto exit;

    if (!GetFunctionPointer(pAppDomain, &vtMethodHandleVal, pMethodAddress))
        goto exit;

    bResult = TRUE;

exit:
    if (bstrClassName) SysFreeString(bstrClassName);
    if (bstrMethodInfoFullName) SysFreeString(bstrMethodInfoFullName);
    if (bstrMethodHandlePropName) SysFreeString(bstrMethodHandlePropName);

    if (pMethods) SafeArrayDestroy(pMethods);

    if (pTargetMethodInfo) pTargetMethodInfo->Release();
    if (pType) pType->Release();
    if (pMethodHandlePropInfo) pMethodHandlePropInfo->Release();
    if (pMethodInfoType) pMethodInfoType->Release();
    if (pReflectionAssembly) pReflectionAssembly->Release();
    if (pAssembly) pAssembly->Release();

    VariantClear(&vtMethodHandleVal);

    return bResult;
}