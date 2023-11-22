#pragma once

#include <AudioClient.h>
#include <mmdeviceapi.h>
#include <mftransform.h>
#include "Wmcodecdsp.h"
#include <atlbase.h>
using namespace ATL;

#include <wrl\implements.h>
#include <wil\com.h>
#include <wil\result.h>

#include <comdef.h>

#include "Common.h"

#define EXIT_ON_ERROR(hres) \
if (FAILED(hres)) \
{\
    _com_error err(hr); LPCTSTR errMsg = err.ErrorMessage(); std::wstring s(errMsg);\
    std::wcout << L"Error in " << __FUNCTIONW__ << ":: " << __LINE__ << ": " << s << std::endl; \
    goto Exit;\
}

/**
* Base class for LoopbackCaptureSync and LoopbackCaptureAsync classes
* 
* Defines common methods and attributes. This class holds the necessary state and methods to resample the captured samples into a sample format compatible with an output client, defined externally.
*/
class LoopbackCaptureBase
{
public:
    // Setters
    void setAudioRenderClient(IAudioRenderClient* rc) { m_OutputRenderClient = rc; }
    void setAudioClient(IAudioClient* ac) { m_OutputAudioClient = ac; }
    void setOutputFormat(WAVEFORMATEXTENSIBLE* wf) { m_pOutputFormat = wf; }
    void setResamplerTransform(IMFTransform* transform) { m_ResamplerTransform = transform; }

    HRESULT initializeMFTResampler(WAVEFORMATEX* inputFmt, WAVEFORMATEXTENSIBLE* outputFmtex);
    // Takes an audio framebuffer and outputs a resampled framebuffer, resampled to the format specified by m_pOutputFormat
    void resampleAudioStream(BYTE* src, BYTE* dst, UINT32 framesAvailable, UINT32 clientFramesAvailable, UINT32& framesWritten);

protected:
    // Output stream to an output endpoint
    IAudioClient* m_OutputAudioClient = NULL;
    // Render client used to play back captured audio
    IAudioRenderClient* m_OutputRenderClient = nullptr;
    // Media Foundations transform for resampling captured samples to a format compatible with the output client (m_pOutputFormat)
    CComPtr<IMFTransform> m_ResamplerTransform = NULL;
    // Sample format compatible with the output client
    WAVEFORMATEXTENSIBLE* m_pOutputFormat = NULL;
    // Sample format of the captured samples
    WAVEFORMATEX m_CaptureFormat{};
    wil::com_ptr_nothrow<IAudioClient> m_AudioClient;
    wil::com_ptr_nothrow<IAudioCaptureClient> m_AudioCaptureClient;
};