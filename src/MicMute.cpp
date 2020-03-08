// MicMute.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <stdio.h>
#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <Functiondiscoverykeys_devpkey.h>
#include "MicMute.h"

#define EXIT_ON_ERROR(hres)  \
              if (FAILED(hres)) { goto Exit; }
#define SAFE_RELEASE(punk)  \
              if ((punk) != NULL)  \
                { (punk)->Release(); (punk) = NULL; }

const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);

void PrintEndpointNames()
{
    HRESULT hr = S_OK;
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDeviceCollection* pCollection = NULL;
    IMMDevice* pEndpoint = NULL;
    IPropertyStore* pProps = NULL;
    LPWSTR pwszID = NULL;

    hr = CoCreateInstance(
        CLSID_MMDeviceEnumerator, NULL,
        CLSCTX_ALL, IID_IMMDeviceEnumerator,
        (void**)&pEnumerator);
    EXIT_ON_ERROR(hr)

        hr = pEnumerator->EnumAudioEndpoints(
            eCapture, DEVICE_STATE_ACTIVE,
            &pCollection);
    EXIT_ON_ERROR(hr)

        UINT  count;
    hr = pCollection->GetCount(&count);
    EXIT_ON_ERROR(hr)

        if (count == 0)
        {
            printf("No endpoints found.\n");
        }

    // Each loop prints the name of an endpoint device.
    for (ULONG i = 0; i < count; i++)
    {
        // Get pointer to endpoint number i.
        hr = pCollection->Item(i, &pEndpoint);
        EXIT_ON_ERROR(hr)

            // Get the endpoint ID string.
            hr = pEndpoint->GetId(&pwszID);
        EXIT_ON_ERROR(hr)

            hr = pEndpoint->OpenPropertyStore(
                STGM_READ, &pProps);
        EXIT_ON_ERROR(hr)

            PROPVARIANT varName;
        // Initialize container for property value.
        PropVariantInit(&varName);

        // Get the endpoint's friendly-name property.
        hr = pProps->GetValue(
            PKEY_Device_FriendlyName, &varName);
        EXIT_ON_ERROR(hr)

            // Print endpoint friendly name and endpoint ID.
            printf("Endpoint %d: \"%S\" (%S)\n",
                i, varName.pwszVal, pwszID);

        CoTaskMemFree(pwszID);
        pwszID = NULL;
        PropVariantClear(&varName);
        SAFE_RELEASE(pProps)
            SAFE_RELEASE(pEndpoint)
    }
    SAFE_RELEASE(pEnumerator)
        SAFE_RELEASE(pCollection)
        return;

Exit:
    printf("Error!\n");
    CoTaskMemFree(pwszID);
    SAFE_RELEASE(pEnumerator)
        SAFE_RELEASE(pCollection)
        SAFE_RELEASE(pEndpoint)
        SAFE_RELEASE(pProps)
}

int GetCaptureId()
{
    HRESULT hr;

    IMMDeviceEnumerator* deviceEnumerator = NULL;
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_INPROC_SERVER, __uuidof(IMMDeviceEnumerator), (LPVOID*)&deviceEnumerator);
    if (FAILED(hr)) {
        std::cout << "CoCreateInstance failed, hr = " << std::hex << hr;
        return 1;
    }

    IMMDevice* defaultDevice = NULL;

    hr = deviceEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &defaultDevice);
    deviceEnumerator->Release();
    deviceEnumerator = NULL;

    if (FAILED(hr)) {
        std::cout << "GetDefaultAudioEndpoint failed, hr = " << std::hex << hr;
        return 2;
    }

    WCHAR* deviceId;
    hr = defaultDevice->GetId(&deviceId);
    if (FAILED(hr)) {
        std::cout << "GetId failed, hr = " << std::hex << hr;
        return 3;
    }

    std::cout << "Default device id: " << deviceId << std::endl;
    CoTaskMemFree(deviceId);

    IAudioEndpointVolume* endpointVolume = NULL;
    hr = defaultDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_INPROC_SERVER, NULL, (LPVOID*)&endpointVolume);
    defaultDevice->Release();
    defaultDevice = NULL;

    BOOL isMuted;
    hr = endpointVolume->GetMute(&isMuted);
    if (FAILED(hr)) {
        std::cout << "GetId failed, hr = " << std::hex << hr;
        return 4;
    }

    std::cout << "Is muted? " << isMuted << std::endl;

    hr = endpointVolume->SetMute(!isMuted, NULL);
    if (FAILED(hr)) {
        std::cout << "SetMute failed, hr = " << std::hex << hr;
        return 5;
    }

    endpointVolume->Release();
    endpointVolume = NULL;

    return 0;
}

int main()
{
    std::cout << "Hello World!\n";

    HRESULT hr;
    hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        std::cout << "CoInitialize failed, hr = " << std::hex << hr;
        return 1;
    }
    // PrintEndpointNames();
    GetCaptureId();

    CoUninitialize();

    return 0;
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
