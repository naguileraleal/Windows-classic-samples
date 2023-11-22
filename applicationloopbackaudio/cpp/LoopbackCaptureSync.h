#pragma once

#include "LoopbackCaptureBase.h"

#include <AudioClient.h>
#include <mmdeviceapi.h>
#include <mftransform.h>
#include "Wmcodecdsp.h"
#include <atlbase.h>
using namespace ATL;

#include <wrl\implements.h>
#include <wil\com.h>
#include <wil\result.h>

#include "Common.h"

using namespace Microsoft::WRL;

class LoopbackCaptureSync :
    public RuntimeClass< RuntimeClassFlags< ClassicCom >, FtmBase, IActivateAudioInterfaceCompletionHandler >, public LoopbackCaptureBase
{
public:
    LoopbackCaptureSync() = default;
    ~LoopbackCaptureSync();

    // IActivateAudioInterfaceCompletionHandler
    STDMETHOD(ActivateCompleted)(IActivateAudioInterfaceAsyncOperation* operation);

    static DWORD WINAPI ThreadProc(LPVOID lpParam)
    {
        LoopbackCaptureSync* This = (LoopbackCaptureSync*)lpParam;
        return This->CaptureThread();
    }

    HRESULT CaptureThread();
    HRESULT StartCapture(DWORD processId, bool includeProcessTree, PCWSTR outputFileName);
    HRESULT StopCapture();

private:
    // NB: All states >= Initialized will allow some methods
    // to be called successfully on the Audio Client
    enum class DeviceState
    {
        Uninitialized,
        Error,
        Initialized,
        Starting,
        Capturing,
        Stopping,
        Stopped,
    };

    HRESULT InitializeLoopbackCapture();
    HRESULT CreateWAVFile();
    HRESULT FixWAVHeader();

    HRESULT ActivateAudioInterface(DWORD processId, bool includeProcessTree);

    HRESULT SetDeviceStateErrorIfFailed(HRESULT hr);

    IAudioClient2* m_AudioClient2;
    UINT32 m_BufferFrames = 0;
    wil::com_ptr_nothrow<IMFAsyncResult> m_SampleReadyAsyncResult;

    wil::unique_event_nothrow m_SampleReadyEvent;
    MFWORKITEM_KEY m_SampleReadyKey = 0;
    wil::unique_hfile m_hFile;
    wil::critical_section m_CritSec;
    DWORD m_dwQueueID = 0;
    DWORD m_cbHeaderSize = 0;
    DWORD m_cbDataSize = 0;

    // These two members are used to communicate between the main thread
    // and the ActivateCompleted callback.
    PCWSTR m_outputFileName = nullptr;
    HRESULT m_activateResult = E_UNEXPECTED;

    DeviceState m_DeviceState{ DeviceState::Uninitialized };
    wil::unique_event_nothrow m_hActivateCompleted;
    wil::unique_event_nothrow m_hCaptureStopped;

    WAVEFORMATEXTENSIBLE* m_pOutputFormat = NULL;
    bool m_bAudioStreamStarted = false;
    UINT64 m_u64QPCPositionPrev = 0;
    IAudioClock* m_pAudioClock = NULL;
    bool m_bStopCapture = false;

    HANDLE m_hCaptureThread = NULL;
    DWORD m_dwThreadId = 0;
};
