// ApplicationLoopback.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <Windows.h>
#include <iostream>
#include "LoopbackCapture.h"
#include "LoopbackCaptureSync.h"

#include <comdef.h>

void usage()
{
    std::wcout <<
        L"Usage: ApplicationLoopback <pid> <includetree|excludetree> <outputfilename> <endpointname> <Sync|Async>\n"
        L"\n"
        L"<pid> is the process ID to capture or exclude from capture\n"
        L"includetree includes audio from that process and its child processes\n"
        L"excludetree includes audio from all processes except that process and its child processes\n"
        L"<outputfilename> is the WAV file to receive the captured audio (10 seconds)\n"
        L"<endpointname> is a substring contained in the friendly name of the audio endpoint where the captured audio will be streamed to\n"
        L"<Sync|Async> use synchronic or asynchronic loopbac capture\n"
        L"\n"
        L"Examples:\n"
        L"\n"
        L"ApplicationLoopback 1234 includetree CapturedAudio.wav Speakers Sync\n"
        L"\n"
        L"  Captures audio from process 1234 and its children, sends it to an audio endpoint that contains the word \"Speakers\" in its name, if it exists. Uses synchronic loopback capture class\n";
}

// REFERENCE_TIME time units per second and per millisecond
#define REFTIMES_PER_SEC  10000000
#define REFTIMES_PER_MILLISEC  10000

#define SAFE_RELEASE(punk)  \
              if ((punk) != NULL)  \
                { (punk)->Release(); (punk) = NULL; }

const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
const IID IID_IAudioClient = __uuidof(IAudioClient);
const IID IID_IAudioRenderClient = __uuidof(IAudioRenderClient);
const CLSID CLSID_CResamplerMediaObject = __uuidof(CResamplerMediaObject);

REFERENCE_TIME hnsRequestedDuration = 0;
REFERENCE_TIME hnsActualDuration;
IMMDeviceEnumerator* pEnumerator = NULL;
IMMDevice* pDevice = NULL;
IAudioClient* pAudioClient = NULL;
IAudioRenderClient* pRenderClient = NULL;
WAVEFORMATEX* pwfx = NULL;
UINT32 bufferFrameCount;
UINT32 numFramesAvailable;
UINT32 numFramesPadding;
BYTE* pData;
DWORD flags = 0;

/**
* Returns a string representation of the given WAVEFORMATEX.wFormatTag.
*/
const char* wFormatTagToString(WORD pFormatTag)
{
    switch (pFormatTag)
    {
        case WAVE_FORMAT_PCM:
            return "WAVE_FORMAT_PCM";
        case WAVE_FORMAT_IEEE_FLOAT:
        return "WAVE_FORMAT_IEEE_FLOAT";
        case WAVE_FORMAT_DRM:
            return "WAVE_FORMAT_DRM";
        case WAVE_FORMAT_ALAW:
            return "WAVE_FORMAT_ALAW";
        case WAVE_FORMAT_MULAW:
            return "WAVE_FORMAT_MULAW";
        case WAVE_FORMAT_ADPCM:
            return "WAVE_FORMAT_ADPCM";
        case WAVE_FORMAT_EXTENSIBLE:
            return "WAVE_FORMAT_EXTENSIBLE";
        case WAVE_FORMAT_WMAUDIO2:
            return "WAVE_FORMAT_WMAUDIO2";
        default:
            return "Unknown";
    }
}

/**
* Returns the index of the first endpoint which has the given substring in its friendly name.
*/
int getEndpointFromFriendlyName(IMMDeviceCollection* pCollection, std::wstring endpointFriendlyName)
{
    HRESULT hr = S_OK;
    int ret = -1;
    IMMDevice* pEndpoint = NULL;
    LPWSTR pwszID = NULL;
    IPropertyStore *pProps = NULL;

    static PROPERTYKEY key;

    GUID IDevice_FriendlyName = { 0xa45c254e, 0xdf1c, 0x4efd, { 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0 } };
    key.pid = 14;
    key.fmtid = IDevice_FriendlyName;

    UINT  count;
    hr = pCollection->GetCount(&count);

    if (count == 0)
    {
        printf("No endpoints found.\n");
    }
    else
    {
        std::wcout << "Found " << count << " endpoints. Searching for " << endpointFriendlyName << std::endl;
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

        hr = pEndpoint->OpenPropertyStore(STGM_READ, &pProps);
        EXIT_ON_ERROR(hr)

        PROPVARIANT varName;
        // Initialize container for property value.
        PropVariantInit(&varName);

        // Get the endpoint's friendly-name property.
        hr = pProps->GetValue(key, &varName);
        EXIT_ON_ERROR(hr)

        // GetValue succeeds and returns S_OK if PKEY_Device_FriendlyName is not found.
        // In this case vartName.vt is set to VT_EMPTY.      
        if (varName.vt != VT_EMPTY)
        {
            // Print endpoint friendly name and endpoint ID.
            printf("Endpoint %d: \"%S\" (%S)\n", i, varName.pwszVal, pwszID);
            std::wstring thisEndpointFriendlyName = varName.pwszVal;
            if (thisEndpointFriendlyName.find(endpointFriendlyName) != std::wstring::npos)
            {
                ret = i;
            }
        }

        CoTaskMemFree(pwszID);
        pwszID = NULL;
        PropVariantClear(&varName);
        SAFE_RELEASE(pProps)
        SAFE_RELEASE(pEndpoint)
    }

    return ret;

Exit:
    return -1;
}

HRESULT calculateMixFormatType(WAVEFORMATEX* _mixFormat)
{
    if (_mixFormat->wFormatTag == WAVE_FORMAT_PCM ||
        _mixFormat->wFormatTag == WAVE_FORMAT_EXTENSIBLE &&
        reinterpret_cast<WAVEFORMATEXTENSIBLE*>(_mixFormat)->SubFormat == KSDATAFORMAT_SUBTYPE_PCM)
    {
        if (_mixFormat->wBitsPerSample == 16)
        {
            printf("PCM 16-bit sample type\n");
        }
        else if (_mixFormat->wBitsPerSample == 24)
        {
            printf("PCM 24-bit sample type\n");
        }
        else if (_mixFormat->wBitsPerSample == 32)
        {
            printf("PCM 32-bit sample type\n");
        }
        else
        {
            printf("Unknown PCM integer sample type\n");
            return E_UNEXPECTED;
        }
    }
    else if (_mixFormat->wFormatTag == WAVE_FORMAT_IEEE_FLOAT || (_mixFormat->wFormatTag == WAVE_FORMAT_EXTENSIBLE && reinterpret_cast<WAVEFORMATEXTENSIBLE*>(_mixFormat)->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT))
    {
        if (_mixFormat->wBitsPerSample == 32)
        {
            printf("Float 32-bit sample type\n");
        }
        else if (_mixFormat->wBitsPerSample == 64)
        {
            printf("Float 64-bit sample type\n");
        }
    }
    else
    {
        printf("unrecognized device format.\n");
        return E_UNEXPECTED;
    }
    return S_OK;
}

/**
* Initializes an audio client to receive the captured stream.
*/
HRESULT initializeOutputClient(LoopbackCaptureBase* capturer, PCWSTR friendlyName)
{
    HRESULT hr;
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDevice* pDevice = NULL;
    IMMDeviceCollection* pCollection = NULL;
    IAudioClient* pAudioClient = NULL;
    IAudioRenderClient* pRenderClient = NULL;
    WAVEFORMATEXTENSIBLE* pDesiredFormat = (WAVEFORMATEXTENSIBLE*)CoTaskMemAlloc(sizeof(WAVEFORMATEXTENSIBLE));
    WAVEFORMATEXTENSIBLE* pClientFormat = (WAVEFORMATEXTENSIBLE*)CoTaskMemAlloc(sizeof(WAVEFORMATEXTENSIBLE));
    WAVEFORMATEXTENSIBLE* pClosestMatch = NULL;
    UINT32 bufferFrameCount;
    UINT32 numFramesAvailable;
    UINT32 numFramesPadding;
    BYTE* pData;
    DWORD flags = 0;
    int i = 0;

    pDesiredFormat->Format.wFormatTag = WAVE_FORMAT_PCM;
    pDesiredFormat->Format.nChannels = 2;
    pDesiredFormat->Format.nSamplesPerSec = 44100;
    pDesiredFormat->Format.wBitsPerSample = 16;
    pDesiredFormat->Format.nBlockAlign = (pDesiredFormat->Format.wBitsPerSample / 8) * pDesiredFormat->Format.nChannels;
    pDesiredFormat->Format.nAvgBytesPerSec = pDesiredFormat->Format.nSamplesPerSec * pDesiredFormat->Format.nBlockAlign;
    pDesiredFormat->Format.cbSize = 0;
    std::cout << "Desired mix format:" << std::endl;
    std::cout << "  wFormatTags    : 0x" << std::hex << pDesiredFormat->Format.wFormatTag << std::dec << std::endl;
    std::cout << "  nChannels      : " << pDesiredFormat->Format.nChannels << std::endl;
    std::cout << "  nSamplesPerSec : " << pDesiredFormat->Format.nSamplesPerSec << std::endl;
    std::cout << "  nAvgBytesPerSec: " << pDesiredFormat->Format.nAvgBytesPerSec << std::endl;
    std::cout << "  nBlockAlign    : " << pDesiredFormat->Format.nBlockAlign << std::endl;
    std::cout << "  wBitsPerSample : " << pDesiredFormat->Format.wBitsPerSample << std::endl;
    std::cout << "  cbSize         : " << pDesiredFormat->Format.cbSize << std::endl;

    hr = CoInitialize(NULL);
    EXIT_ON_ERROR(hr)

    hr = CoCreateInstance(
        CLSID_MMDeviceEnumerator, NULL,
        CLSCTX_ALL, IID_IMMDeviceEnumerator,
        (void**)&pEnumerator);
    EXIT_ON_ERROR(hr)

    hr = pEnumerator->EnumAudioEndpoints(
        eRender, DEVICE_STATE_ACTIVE,
        &pCollection);
    EXIT_ON_ERROR(hr)

    // Use a specific Audio Endpoint instead of using the default one
    /*
    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
    EXIT_ON_ERROR(hr)
    */
    // e.g: Realtek HD Audio 2nd output (Realtek(R) Audio Codec with THX Spatial Audio)
    i = getEndpointFromFriendlyName(pCollection, friendlyName);

    if (i == -1)
    {
        std::cout << "Endpoint not found" << std::endl;
    }
    else
    {
        hr = pCollection->Item(i, &pDevice);
        EXIT_ON_ERROR(hr)

        hr = pDevice->Activate(IID_IAudioClient, CLSCTX_ALL, NULL, (void**)&pAudioClient);
        EXIT_ON_ERROR(hr)

        hr = pAudioClient->GetMixFormat(reinterpret_cast<WAVEFORMATEX**>(&pClientFormat));
        EXIT_ON_ERROR(hr)
        calculateMixFormatType(reinterpret_cast<WAVEFORMATEX*>(pClientFormat));

        std::cout << "Endpoint's mix format:" << std::endl;
        std::cout << "  wFormatTags    : " << wFormatTagToString(pClientFormat->Format.wFormatTag) << std::dec << std::endl;
        std::cout << "  nChannels      : " << pClientFormat->Format.nChannels << std::endl;
        std::cout << "  nSamplesPerSec : " << pClientFormat->Format.nSamplesPerSec << std::endl;
        std::cout << "  nAvgBytesPerSec: " << pClientFormat->Format.nAvgBytesPerSec << std::endl;
        std::cout << "  nBlockAlign    : " << pClientFormat->Format.nBlockAlign << std::endl;
        std::cout << "  wBitsPerSample : " << pClientFormat->Format.wBitsPerSample << std::endl;
        std::cout << "  cbSize         : " << pClientFormat->Format.cbSize << std::endl;

        hr = pAudioClient->IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, &pDesiredFormat->Format, reinterpret_cast<WAVEFORMATEX**>(&pClosestMatch));

        if (hr == AUDCLNT_E_UNSUPPORTED_FORMAT)
        {
            std::cout << "Desired format is unsupported\n";
            hr = pAudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED, 0, 0, 0, &pClientFormat->Format, NULL);
            EXIT_ON_ERROR(hr)
            capturer->setOutputFormat(pClientFormat);

        }
        else if (hr == S_OK && pClosestMatch == NULL)
        {
                std::cout << "Desired format is supported\n";
                hr = pAudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED, 0, 0, 0, &pDesiredFormat->Format, NULL);
                _com_error err(hr);
                LPCTSTR errMsg = err.ErrorMessage();
                std::cout << "Error: " << errMsg << std::endl;
                EXIT_ON_ERROR(hr)
                capturer->setOutputFormat(pDesiredFormat);
        }
        else if (hr == S_FALSE)
        {
            std::cout << "Desired format is not supported, but closest match is\n";
            calculateMixFormatType(reinterpret_cast<WAVEFORMATEX*>(pClosestMatch));
            std::cout << "Endpoint's closest match to desired sample format:" << std::endl;
            std::cout << "  wFormatTags    : " << wFormatTagToString(pClosestMatch->Format.wFormatTag) << std::endl;
            std::cout << "  nChannels      : " << pClosestMatch->Format.nChannels << std::endl;
            std::cout << "  nSamplesPerSec : " << pClosestMatch->Format.nSamplesPerSec << std::endl;
            std::cout << "  nAvgBytesPerSec: " << pClosestMatch->Format.nAvgBytesPerSec << std::endl;
            std::cout << "  nBlockAlign    : " << pClosestMatch->Format.nBlockAlign << std::endl;
            std::cout << "  wBitsPerSample : " << pClosestMatch->Format.wBitsPerSample << std::endl;
            std::cout << "  cbSize         : " << pClosestMatch->Format.cbSize << std::endl;
            hr = pAudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED, 0, 0, 0, &pClosestMatch->Format, NULL);
            EXIT_ON_ERROR(hr)
            capturer->setOutputFormat(pClosestMatch);
        }

        hr = pAudioClient->GetService(IID_IAudioRenderClient, (void**)&pRenderClient);
        EXIT_ON_ERROR(hr)

        capturer->setAudioRenderClient(pRenderClient);
        capturer->setAudioClient(pAudioClient);
    }

    return hr;

Exit:
    CoTaskMemFree(pwfx);
    SAFE_RELEASE(pEnumerator)
    SAFE_RELEASE(pDevice)
    SAFE_RELEASE(pAudioClient)
    SAFE_RELEASE(pRenderClient)

    return hr;
}

void loopbackCaptureSync(DWORD processId, bool includeProcessTree, PCWSTR outputFile, PCWSTR outputFriendlyName)
{
    LoopbackCaptureSync loopbackCapture;
    initializeOutputClient(&loopbackCapture, outputFriendlyName);

    HRESULT hr = loopbackCapture.StartCapture(processId, includeProcessTree, outputFile);
    if (FAILED(hr))
    {
        wil::unique_hlocal_string message;
        FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_ALLOCATE_BUFFER, nullptr, hr,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (PWSTR)&message, 0, nullptr);
        std::wcout << L"Failed to start capture\n0x" << std::hex << hr << L": " << message.get() << L"\n";
    }
    else
    {
        std::wcout << L"Capturing 1000 seconds of audio." << std::endl;
        Sleep(1000000);

        loopbackCapture.StopCapture();

        std::wcout << L"Finished.\n";
    }
}

void loopbackCaptureAsync(DWORD processId, bool includeProcessTree, PCWSTR outputFile, PCWSTR outputFriendlyName)
{
    CLoopbackCapture loopbackCapture;
    initializeOutputClient(&loopbackCapture, outputFriendlyName);

    HRESULT hr = loopbackCapture.StartCaptureAsync(processId, includeProcessTree, outputFile);
    if (FAILED(hr))
    {
        wil::unique_hlocal_string message;
        FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_ALLOCATE_BUFFER, nullptr, hr,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (PWSTR)&message, 0, nullptr);
        std::wcout << L"Failed to start capture\n0x" << std::hex << hr << L": " << message.get() << L"\n";
    }
    else
    {
        std::wcout << L"Capturing 1000 seconds of audio." << std::endl;
        Sleep(1000000);

        loopbackCapture.StopCaptureAsync();

        std::wcout << L"Finished.\n";
    }
}

int wmain(int argc, wchar_t* argv[])
{
    if (argc != 6)
    {
        usage();
        return 0;
    }

    DWORD processId = wcstoul(argv[1], nullptr, 0);
    if (processId == 0)
    {
        usage();
        return 0;
    }

    bool includeProcessTree;
    if (wcscmp(argv[2], L"includetree") == 0)
    {
        includeProcessTree = true;
    }
    else if (wcscmp(argv[2], L"excludetree") == 0)
    {
        includeProcessTree = false;
    }
    else
    {
        usage();
        return 0;
    }

    PCWSTR outputFile = argv[3];

    PCWSTR outputFriendlyName = argv[4];
    
    // Synchronous or asynchronous mode
    PCWSTR mode = argv[5];

    if (wcscmp(mode, L"Sync") == 0)
    {
        loopbackCaptureSync(processId, includeProcessTree, outputFile, outputFriendlyName);
    }
    else if (wcscmp(mode, L"Async") == 0)
    {
        loopbackCaptureAsync(processId, includeProcessTree, outputFile, outputFriendlyName);
    }


    CoUninitialize();

    return 0;
}
