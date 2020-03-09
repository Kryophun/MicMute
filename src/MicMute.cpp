// MicMute.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <atlbase.h>
#include <wil\Result.h>
#include "MicMute.h"

void ToggleMuteOnDefaultCaptureDevice()
{
    CComPtr<IMMDeviceEnumerator> deviceEnumerator;
    THROW_IF_FAILED_MSG(deviceEnumerator.CoCreateInstance(__uuidof(MMDeviceEnumerator)), "Could not create device enumerator");

    CComPtr<IMMDevice> defaultDevice;
    THROW_IF_FAILED_MSG(deviceEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &defaultDevice), "Could not get default audio capture endpoint");

    CComPtr<IAudioEndpointVolume> endpointVolume;
    THROW_IF_FAILED_MSG(defaultDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_INPROC_SERVER, NULL, (LPVOID*)&endpointVolume), "Could not get audio endpoint volume");

    BOOL isMuted;
    THROW_IF_FAILED_MSG(endpointVolume->GetMute(&isMuted), "Could not get mute state of audio endpoint");

    std::cout << "Device is currently " << (isMuted ? "muted. Turning mute off..." : "NOT muted. Muting device...") << std::endl;

    THROW_IF_FAILED_MSG(endpointVolume->SetMute(!isMuted, NULL), "Could not set mute state");
}

int main()
{
    std::cout << "Hello World!\n";
    int returnCode = 0;

    try {
        THROW_IF_FAILED_MSG(CoInitialize(NULL), "CoInitialize Failed");
        ToggleMuteOnDefaultCaptureDevice();
    }
    catch (wil::ResultException& e) {
        const wil::FailureInfo failure = e.GetFailureInfo();
        std::cout << "An error occurred: (" << std::hex << failure.hr << "), message = " << failure.pszMessage << std::endl;
        returnCode = 1;
    }

    CoUninitialize();

    return returnCode;
}

// Perhaps this will be helpful in the future

//#define EXIT_ON_ERROR(hres)  \
//              if (FAILED(hres)) { goto Exit; }
//#define SAFE_RELEASE(punk)  \
//              if ((punk) != NULL)  \
//                { (punk)->Release(); (punk) = NULL; }
//
//const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
//const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);

//void PrintEndpointNames()
//{
//    HRESULT hr = S_OK;
//    IMMDeviceEnumerator* pEnumerator = NULL;
//    IMMDeviceCollection* pCollection = NULL;
//    IMMDevice* pEndpoint = NULL;
//    IPropertyStore* pProps = NULL;
//    LPWSTR pwszID = NULL;
//
//    hr = CoCreateInstance(
//        CLSID_MMDeviceEnumerator, NULL,
//        CLSCTX_ALL, IID_IMMDeviceEnumerator,
//        (void**)&pEnumerator);
//    EXIT_ON_ERROR(hr)
//
//        hr = pEnumerator->EnumAudioEndpoints(
//            eCapture, DEVICE_STATE_ACTIVE,
//            &pCollection);
//    EXIT_ON_ERROR(hr)
//
//        UINT  count;
//    hr = pCollection->GetCount(&count);
//    EXIT_ON_ERROR(hr)
//
//        if (count == 0)
//        {
//            printf("No endpoints found.\n");
//        }
//
//    // Each loop prints the name of an endpoint device.
//    for (ULONG i = 0; i < count; i++)
//    {
//        // Get pointer to endpoint number i.
//        hr = pCollection->Item(i, &pEndpoint);
//        EXIT_ON_ERROR(hr)
//
//            // Get the endpoint ID string.
//            hr = pEndpoint->GetId(&pwszID);
//        EXIT_ON_ERROR(hr)
//
//            hr = pEndpoint->OpenPropertyStore(
//                STGM_READ, &pProps);
//        EXIT_ON_ERROR(hr)
//
//            PROPVARIANT varName;
//        // Initialize container for property value.
//        PropVariantInit(&varName);
//
//        // Get the endpoint's friendly-name property.
//        hr = pProps->GetValue(
//            PKEY_Device_FriendlyName, &varName);
//        EXIT_ON_ERROR(hr)
//
//            // Print endpoint friendly name and endpoint ID.
//            printf("Endpoint %d: \"%S\" (%S)\n",
//                i, varName.pwszVal, pwszID);
//
//        if (wcsstr(varName.bstrVal, L"C615") != NULL)
//        {
//            printf("HUZZAH!!");
//            CComPtr<IAudioEndpointVolume> endpointVolume;
//            hr = pEndpoint->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_INPROC_SERVER, NULL, (LPVOID*)&endpointVolume);
//
//            BOOL isMuted;
//            hr = endpointVolume->GetMute(&isMuted);
//            if (FAILED(hr)) {
//                std::cout << "GetId failed, hr = " << std::hex << hr;
//                return;
//            }
//
//            std::cout << "Is muted? " << isMuted << std::endl;
//
//            hr = endpointVolume->SetMute(!isMuted, NULL);
//            if (FAILED(hr)) {
//                std::cout << "SetMute failed, hr = " << std::hex << hr;
//                return;
//            }
//        }
//
//        CoTaskMemFree(pwszID);
//        pwszID = NULL;
//        PropVariantClear(&varName);
//        SAFE_RELEASE(pProps)
//            SAFE_RELEASE(pEndpoint)
//    }
//    SAFE_RELEASE(pEnumerator)
//        SAFE_RELEASE(pCollection)
//        return;
//
//Exit:
//    printf("Error!\n");
//    CoTaskMemFree(pwszID);
//    SAFE_RELEASE(pEnumerator)
//        SAFE_RELEASE(pCollection)
//        SAFE_RELEASE(pEndpoint)
//        SAFE_RELEASE(pProps)
//}
