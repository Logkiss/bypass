#include<bits/stdc++.h>
#include <taskschd.h>
#include <comutil.h>

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")

void Cleanup(ITaskService* pService, ITaskFolder* pRootFolder, ITaskDefinition* pTask) {
    if (pRootFolder) pRootFolder->Release();
    if (pTask) pTask->Release();
    if (pService) pService->Release();
    CoUninitialize();
}

int main() {
    // 初始化 COM 库
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        return 1;
    }

    ITaskService* pService = nullptr;
    hr = CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService);
    if (FAILED(hr)) {
        CoUninitialize();
        return 1;
    }

    hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
    if (FAILED(hr)) {
        Cleanup(pService, nullptr, nullptr);
        return 1;
    }

    ITaskFolder* pRootFolder = nullptr;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
    if (FAILED(hr)) {
        Cleanup(pService, nullptr, nullptr);
        return 1;
    }

    ITaskDefinition* pTask = nullptr;
    hr = pService->NewTask(0, &pTask);
    if (FAILED(hr)) {
        Cleanup(pService, pRootFolder, nullptr);
        return 1;
    }

    ITaskSettings* pSettings = nullptr;
    hr = pTask->get_Settings(&pSettings);
    if (SUCCEEDED(hr)) {
        pSettings->put_StartWhenAvailable(VARIANT_TRUE);
        pSettings->Release();
    }

    IRegistrationInfo* pInfo = nullptr;
    hr = pTask->get_RegistrationInfo(&pInfo);
    if (SUCCEEDED(hr)) {
        pInfo->put_Author(_bstr_t(L"Microsoft"));
        pInfo->Release();
    }

    ITriggerCollection* pTriggers = nullptr;
    hr = pTask->get_Triggers(&pTriggers);
    if (SUCCEEDED(hr)) {
        ITrigger* pTrigger = nullptr;
        hr = pTriggers->Create(TASK_TRIGGER_LOGON, &pTrigger);
        if (SUCCEEDED(hr)) {
            pTrigger->Release();
        }
        pTriggers->Release();
    }

    IActionCollection* pActions = nullptr;
    hr = pTask->get_Actions(&pActions);
    if (SUCCEEDED(hr)) {
        IAction* pAction = nullptr;
        hr = pActions->Create(TASK_ACTION_EXEC, &pAction);
        if (SUCCEEDED(hr)) {
            IExecAction* pExecAction = nullptr;
            hr = pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);
            if (SUCCEEDED(hr)) {
                pExecAction->put_Path(_bstr_t(L"C:\\test.exe"));
                pExecAction->Release();
            }
            pAction->Release();
        }
        pActions->Release();
    }

    IRegisteredTask* pRegisteredTask = nullptr;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(L"Test Task"),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(L"system"),
        _variant_t(),
        TASK_LOGON_GROUP,
        _variant_t(L""),
        &pRegisteredTask
    );

    if (FAILED(hr)) {
    } else {

        pRegisteredTask->Release();
    }

    Cleanup(pService, pRootFolder, pTask);
    return 0;
}
