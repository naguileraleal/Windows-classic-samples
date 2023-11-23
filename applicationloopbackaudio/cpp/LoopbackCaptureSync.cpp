#include <wchar.h>
#include <iostream>
#include <audioclientactivationparams.h>

#include <comdef.h>

#include "LoopbackCaptureSync.h"

HRESULT LoopbackCaptureSync::SetDeviceStateErrorIfFailed(HRESULT hr)
{
    if (FAILED(hr))
    {
        m_DeviceState = DeviceState::Error;
    }
    return hr;
}

HRESULT LoopbackCaptureSync::InitializeLoopbackCapture()
{
    // Initialize MF
    RETURN_IF_FAILED(MFStartup(MF_VERSION, MFSTARTUP_LITE));

    // Create the completion event as auto-reset
    RETURN_IF_FAILED(m_hActivateCompleted.create(wil::EventOptions::None));

    // Create the capture-stopped event as auto-reset
    RETURN_IF_FAILED(m_hCaptureStopped.create(wil::EventOptions::None));

    return S_OK;
}

LoopbackCaptureSync::~LoopbackCaptureSync()
{
    if (m_dwQueueID != 0)
    {
        MFUnlockWorkQueue(m_dwQueueID);
    }
}

HRESULT LoopbackCaptureSync::ActivateAudioInterface(DWORD processId, bool includeProcessTree)
{
    return SetDeviceStateErrorIfFailed([&]() -> HRESULT
        {
            AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = {};
            audioclientActivationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
            audioclientActivationParams.ProcessLoopbackParams.ProcessLoopbackMode = includeProcessTree ?
                PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE : PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE;
            audioclientActivationParams.ProcessLoopbackParams.TargetProcessId = processId;

            PROPVARIANT activateParams = {};
            activateParams.vt = VT_BLOB;
            activateParams.blob.cbSize = sizeof(audioclientActivationParams);
            activateParams.blob.pBlobData = (BYTE*)&audioclientActivationParams;

            wil::com_ptr_nothrow<IActivateAudioInterfaceAsyncOperation> asyncOp;
            // Cant request IAudioClient2 or IAudioClient3 bc it returns with error...
            RETURN_IF_FAILED(ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &activateParams, this, &asyncOp));

            // Wait for activation completion
            m_hActivateCompleted.wait();

            return m_activateResult;
        }());
}

static bool compareFormats(WAVEFORMATEX* w1, WAVEFORMATEX* w2)
{
    bool res = true;
    // Compare each field. If they differ in a single value, return false. Else, return true
    res &= (w1->wFormatTag == w2->wFormatTag);
    res &= (w1->nChannels == w2->nChannels);
    res &= (w1->nSamplesPerSec == w2->nSamplesPerSec);
    res &= (w1->wBitsPerSample == w2->wBitsPerSample);
    res &= (w1->nBlockAlign == w2->nBlockAlign);
    res &= (w1->nAvgBytesPerSec == w2->nAvgBytesPerSec);
    res &= (w1->cbSize == w2->cbSize);
    return res;
}

//
//  ActivateCompleted()
//
//  Callback implementation of ActivateAudioInterfaceAsync function.  This will be called on MTA thread
//  when results of the activation are available.
//
HRESULT LoopbackCaptureSync::ActivateCompleted(IActivateAudioInterfaceAsyncOperation* operation)
{
    m_activateResult = SetDeviceStateErrorIfFailed([&]()->HRESULT
        {
            std::cout << "Initializing LoopbackCaptureBase's Audio Client" << std::endl;
            // Check for a successful activation result
            HRESULT hrActivateResult = E_UNEXPECTED;
            wil::com_ptr_nothrow<IUnknown> punkAudioInterface;
            RETURN_IF_FAILED(operation->GetActivateResult(&hrActivateResult, &punkAudioInterface));
            RETURN_IF_FAILED(hrActivateResult);

            // Get the pointer for the Audio Client
            RETURN_IF_FAILED(punkAudioInterface.copy_to(&m_AudioClient));

            // Try to upgrade IAudioClient
            auto hr = m_AudioClient->QueryInterface(__uuidof(IAudioClient2), reinterpret_cast<void**>(&m_AudioClient2));
            if (FAILED(hr))
            {
                // Print human-readable error
                _com_error err(hr);
                LPCTSTR errMsg = err.ErrorMessage();
                std::wstring s(errMsg);
                std::wcout << L"QueryInterface Error: " << s << std::endl;
            }

            // output format will be null if there's no output client
            if (m_pOutputFormat)
            {
                if (compareFormats(&m_CaptureFormat, &m_pOutputFormat->Format))
                {
                    // No resampling needed
                    std::cout << "Capture and output formats are identical. No resampling needed." << std::endl;
                }
                else
                {
                    std::cout << "Capture and output formats are different. Resampling needed." << std::endl;
                    //if (m_pOutputFormat->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT)
                    //{
                    //	std::cout << "Output sample format is IEEE_FLOAT. Performing manual resampling." << std::endl;
                    //}
                    //else
                    {
                        std::cout << "Output sample format is different than IEEE_FLOAT. Initializing Media Foundation resampler." << std::endl;
                        // Resampling is needed
                        RETURN_IF_FAILED(initializeMFTResampler(&m_CaptureFormat, m_pOutputFormat));
                    }
                }
            }

            // Get the device period
            REFERENCE_TIME hnsDefaultDevicePeriod;
            REFERENCE_TIME hnsMinimumDevicePeriod;
            hr = m_AudioClient->GetDevicePeriod(&hnsDefaultDevicePeriod, &hnsMinimumDevicePeriod);
            if (FAILED(hr))
            {
                // Print message from hresult
                _com_error err(hr);
                LPCTSTR errMsg = err.ErrorMessage();
                std::wstring s(errMsg);
                std::wcout << L"GetDevicePeriod error: " << s << std::endl;
            }
            else
            {
                std::cout << "Default device period: " << hnsDefaultDevicePeriod << " 100-nanoseconds" << " Minimum device period: " << hnsMinimumDevicePeriod << " 100-nanoseconds" << std::endl;
            }

            // Initialize the AudioClient in Shared Mode with the user specified buffer
            // Note: It is a lie that the audio client resamples the audio to the output format. It does not even resample it to the sample rate!
            RETURN_IF_FAILED(m_AudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_LOOPBACK,
                0,
                0,
                &m_CaptureFormat,
                nullptr));

            //WAVEFORMATEX* effectiveCaptureFormat = {};
            //hr = m_AudioClient->GetMixFormat(&effectiveCaptureFormat);
            //std::cout << "Capture client effective mix format: ";
            //std::cout << "  wFormatTags    : 0x" << std::hex << effectiveCaptureFormat->wFormatTag << std::dec << std::endl;
            //std::cout << "  nChannels      : " << effectiveCaptureFormat->nChannels << std::endl;
            //std::cout << "  nSamplesPerSec : " << effectiveCaptureFormat->nSamplesPerSec << std::endl;
            //std::cout << "  nAvgBytesPerSec: " << effectiveCaptureFormat->nAvgBytesPerSec << std::endl;
            //std::cout << "  nBlockAlign    : " << effectiveCaptureFormat->nBlockAlign << std::endl;
            //std::cout << "  wBitsPerSample : " << effectiveCaptureFormat->wBitsPerSample << std::endl;
            //std::cout << "  cbSize         : " << effectiveCaptureFormat->cbSize << std::endl;


            // Get the maximum size of the AudioClient Buffer
            RETURN_IF_FAILED(m_AudioClient->GetBufferSize(&m_BufferFrames));
            std::cout << "Buffer size: " << m_BufferFrames << " frames" << std::endl;

            // Get an Audio Clock to retrieve the current position in the stream
            // Note: This is not the same as the system clock!
            hr = m_AudioClient->GetService(IID_PPV_ARGS(&m_pAudioClock));
            if (FAILED(hr))
            {
                // Print message from hresult
                _com_error err(hr);
                LPCTSTR errMsg = err.ErrorMessage();
                std::wstring s(errMsg);
                std::wcout << L"Error while requesting an AudioClock: " << s << std::endl;
            }

            // Get the capture client
            RETURN_IF_FAILED(m_AudioClient->GetService(IID_PPV_ARGS(&m_AudioCaptureClient)));

            // Creates the WAV file.
            RETURN_IF_FAILED(CreateWAVFile());

            // Everything is ready.
            m_DeviceState = DeviceState::Initialized;

            return S_OK;
        }());

    // Let ActivateAudioInterface know that m_activateResult has the result of the activation attempt.
    m_hActivateCompleted.SetEvent();
    return S_OK;
}

//
//  CreateWAVFile()
//
//  Creates a WAV file in music folder
//
HRESULT LoopbackCaptureSync::CreateWAVFile()
{
    return SetDeviceStateErrorIfFailed([&]()->HRESULT
        {
            m_hFile.reset(CreateFile(m_outputFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
            RETURN_LAST_ERROR_IF(!m_hFile);

            // Create and write the WAV header
            if (m_CaptureFormat.wFormatTag == WAVE_FORMAT_EXTENSIBLE)
            {
                // WAVEFORMATEXTENSIBLE does not fit in a .wav file. In this case, generate a file without header.
            }
            else
            {
                // 1. RIFF chunk descriptor
                DWORD header[] = {
                    FCC('RIFF'),        // RIFF header
                    0,                  // Total size of WAV (will be filled in later)
                    FCC('WAVE'),        // WAVE FourCC
                    FCC('fmt '),        // Start of 'fmt ' chunk
                    sizeof(m_CaptureFormat) // Size of fmt chunk
                };
                DWORD dwBytesWritten = 0;
                RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), header, sizeof(header), &dwBytesWritten, NULL));

                m_cbHeaderSize += dwBytesWritten;

                // 2. The fmt sub-chunk
                WI_ASSERT(m_CaptureFormat.cbSize == 0);
                RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_CaptureFormat, sizeof(m_CaptureFormat), &dwBytesWritten, NULL));
                m_cbHeaderSize += dwBytesWritten;

                // 3. The data sub-chunk
                DWORD data[] = { FCC('data'), 0 };  // Start of 'data' chunk
                RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), data, sizeof(data), &dwBytesWritten, NULL));
                m_cbHeaderSize += dwBytesWritten;
            }

            return S_OK;
        }());
}


//
//  FixWAVHeader()
//
//  The size values were not known when we originally wrote the header, so now go through and fix the values
//
HRESULT LoopbackCaptureSync::FixWAVHeader()
{
    // Write the size of the 'data' chunk first
    DWORD dwPtr = SetFilePointer(m_hFile.get(), m_cbHeaderSize - sizeof(DWORD), NULL, FILE_BEGIN);
    RETURN_LAST_ERROR_IF(INVALID_SET_FILE_POINTER == dwPtr);

    DWORD dwBytesWritten = 0;
    RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_cbDataSize, sizeof(DWORD), &dwBytesWritten, NULL));

    // Write the total file size, minus RIFF chunk and size
    // sizeof(DWORD) == sizeof(FOURCC)
    RETURN_LAST_ERROR_IF(INVALID_SET_FILE_POINTER == SetFilePointer(m_hFile.get(), sizeof(DWORD), NULL, FILE_BEGIN));

    DWORD cbTotalSize = m_cbDataSize + m_cbHeaderSize - 8;
    RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &cbTotalSize, sizeof(DWORD), &dwBytesWritten, NULL));

    RETURN_IF_WIN32_BOOL_FALSE(FlushFileBuffers(m_hFile.get()));

    return S_OK;
}

HRESULT LoopbackCaptureSync::StartCapture(DWORD processId, bool includeProcessTree, PCWSTR outputFileName)
{
    m_outputFileName = outputFileName;
    auto resetOutputFileName = wil::scope_exit([&] { m_outputFileName = nullptr; });

    RETURN_IF_FAILED(InitializeLoopbackCapture());
    RETURN_IF_FAILED(ActivateAudioInterface(processId, includeProcessTree));

    // We should be in the initialzied state if this is the first time through getting ready to capture.
    if (m_DeviceState == DeviceState::Initialized)
    {
        m_DeviceState = DeviceState::Starting;

        // Start the capture
        RETURN_IF_FAILED(m_AudioClient->Start());

        // Get the device period
        REFERENCE_TIME streamLatency;
        auto hr = m_AudioClient->GetStreamLatency(&streamLatency);
        if (FAILED(hr))
        {
            // Print message from hresult
            _com_error err(hr);
            LPCTSTR errMsg = err.ErrorMessage();
            std::wstring s(errMsg);
            std::wcout << L"GetStreamLatency error: " << s << std::endl;
        }
        else
        {
            std::cout << "StreamLatency: " << streamLatency << std::endl;
        }

        m_DeviceState = DeviceState::Capturing;
        HANDLE m_hCaptureThread = CreateThread( 
            NULL,                   // default security attributes
            0,                      // use default stack size  
            ThreadProc,       // thread function name
            this,          // argument to thread function 
            0,                      // use default creation flags 
            &m_dwThreadId);   // returns the thread identifier 

        m_hCaptureStopped.wait();
    }

    return S_OK;
}


//
//  StopCapture()
//
HRESULT LoopbackCaptureSync::StopCapture()
{

    HRESULT hr;

    RETURN_HR_IF(E_NOT_VALID_STATE, (m_DeviceState != DeviceState::Capturing) && (m_DeviceState != DeviceState::Error));

    m_DeviceState = DeviceState::Stopping;

    m_AudioClient->Stop();

    if (m_OutputAudioClient != nullptr)
    {
        m_OutputAudioClient->Stop();
    }

    if (m_CaptureFormat.wFormatTag != WAVE_FORMAT_EXTENSIBLE)
    {
        hr = FixWAVHeader();
    }

    // Stop MFTransform
    if (m_ResamplerTransform != nullptr)
    {
        // This should happen right before the last audio frame, and then the resampler function should be called one last time
        hr = m_ResamplerTransform->ProcessMessage(MFT_MESSAGE_NOTIFY_END_OF_STREAM, NULL);
        hr = m_ResamplerTransform->ProcessMessage(MFT_MESSAGE_COMMAND_DRAIN, NULL);

        // Shut down the resampler
        hr = m_ResamplerTransform->ProcessMessage(MFT_MESSAGE_NOTIFY_END_STREAMING, NULL);
        MFShutdown();
    }

    m_DeviceState = DeviceState::Stopped;
    m_hCaptureStopped.SetEvent();

    return S_OK;
}

//
//  CaptureThread()
//
HRESULT LoopbackCaptureSync::CaptureThread()
{
    UINT32 FramesAvailable = 0;
    BYTE* Data = nullptr;
    DWORD dwCaptureFlags;
    // Number of frames since the start of a stream until the first frame of the audio packet
    UINT64 u64DevicePosition = 0;
    // Time at which the first frame of the audio packet was written to the endpoint buffer, in 100-nanosecond units
    UINT64 u64QPCPosition = 0;
    DWORD cbBytesToCapture = 0;
    HRESULT hr = S_OK;

    while (m_DeviceState == DeviceState::Capturing)
    {
        // A word on why we have a loop here;
        // Suppose it has been 10 milliseconds or so since the last time
        // this routine was invoked, and that we're capturing 48000 samples per second.
        //
        // The audio engine can be reasonably expected to have accumulated about that much
        // audio data - that is, about 480 samples.
        //
        // However, the audio engine is free to accumulate this in various ways:
        // a. as a single packet of 480 samples, OR
        // b. as a packet of 80 samples plus a packet of 400 samples, OR
        // c. as 48 packets of 10 samples each.
        //
        // In particular, there is no guarantee that this routine will be
        // run once for each packet.
        //
        // So every time this routine runs, we need to read ALL the packets
        // that are now available;
        //
        // We do this by calling IAudioCaptureClient::GetNextPacketSize
        // over and over again until it indicates there are no more packets remaining.
        while (SUCCEEDED(m_AudioCaptureClient->GetNextPacketSize(&FramesAvailable)) && FramesAvailable > 0)
        {
            cbBytesToCapture = FramesAvailable * m_CaptureFormat.nBlockAlign;

            // WAV files have a 4GB (0xFFFFFFFF) size limit, so likely we have hit that limit when we
            // overflow here.  Time to stop the capture
            if ((m_cbDataSize + cbBytesToCapture) < m_cbDataSize)
            {
                return S_OK;
            }

            // Get sample buffer
            RETURN_IF_FAILED(m_AudioCaptureClient->GetBuffer(&Data, &FramesAvailable, &dwCaptureFlags, &u64DevicePosition, &u64QPCPosition));
            std::cout << "Packet ID: " << u64DevicePosition << "Packet ID HEX: " << std::hex << u64DevicePosition << std::dec << " Timestamp: " << u64QPCPosition << std::endl;

            if (dwCaptureFlags & AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR)
            {
                std::cout << "Timestamp error!" << std::endl;
            }
            else
            {
                UINT32 captureClientCurrentPadding = 0;
                m_AudioClient->GetCurrentPadding(&captureClientCurrentPadding);
                std::cout << "Packet ID: " << u64DevicePosition << " Timestamp: " << u64QPCPosition << " Current Padding: " << captureClientCurrentPadding << " Time since last packet : " << (u64QPCPosition - m_u64QPCPositionPrev) << " 100 - nanoseconds" << std::endl;
                m_u64QPCPositionPrev = u64QPCPosition;
                LARGE_INTEGER frequency, endTime;
                // Ticks per second
                QueryPerformanceFrequency(&frequency);
                QueryPerformanceCounter(&endTime);
                LONGLONG endTimeMicroseconds = (endTime.QuadPart * 1000000) / frequency.QuadPart;
                LONGLONG elapsedTime = endTimeMicroseconds - (u64QPCPosition/10);
                if (elapsedTime > 15000)
                {
                    std::cout << "Time elapsed since the first frame of the audio packet was written: " << elapsedTime << " us" << std::endl;
                    // Release the loopback capture's buffer back
                    hr = m_AudioCaptureClient->ReleaseBuffer(FramesAvailable);
                    RETURN_IF_FAILED(hr);
                    // Discard samples that are older than 10ms
                    while (SUCCEEDED(m_AudioCaptureClient->GetNextPacketSize(&FramesAvailable)) && FramesAvailable > 0)
                    {
                        RETURN_IF_FAILED(m_AudioCaptureClient->GetBuffer(&Data, &FramesAvailable, &dwCaptureFlags, &u64DevicePosition, &u64QPCPosition));
                        hr = m_AudioCaptureClient->ReleaseBuffer(FramesAvailable);
                        RETURN_IF_FAILED(hr);
                    }

                    if (m_OutputAudioClient != nullptr && m_bAudioStreamStarted)
                    {
                        m_OutputAudioClient->Stop();
                        m_OutputAudioClient->Reset();
                        m_OutputAudioClient->Start();
                    }
                    std::cout << "Discarded all late samples" << std::endl;

                    continue;
                }

            }

            // Stream to endpoint
            if (m_OutputAudioClient != nullptr)
            {
                if (!m_bAudioStreamStarted)
                {
                    hr = m_OutputAudioClient->Start();
                    RETURN_IF_FAILED(hr);
                    m_bAudioStreamStarted = true;

                    // Initialize audio resampler transform (if needed)
                    if (m_ResamplerTransform != nullptr)
                    {
                        hr = m_ResamplerTransform->ProcessMessage(MFT_MESSAGE_COMMAND_FLUSH, NULL);
                        hr = m_ResamplerTransform->ProcessMessage(MFT_MESSAGE_NOTIFY_BEGIN_STREAMING, NULL);
                        hr = m_ResamplerTransform->ProcessMessage(MFT_MESSAGE_NOTIFY_START_OF_STREAM, NULL);
                    }
                }

                // See how much buffer space is available.
                UINT32 numFramesPadding = 0;
                auto hr = m_OutputAudioClient->GetCurrentPadding(&numFramesPadding);
                RETURN_IF_FAILED(hr);
                std::cout << "Output client current padding: " << numFramesPadding << std::endl;

                // Get the actual size of the allocated buffer.
                UINT32 bufferFrameCount = 0;
                hr = m_OutputAudioClient->GetBufferSize(&bufferFrameCount);
                RETURN_IF_FAILED(hr);

                // Space available in the render client's buffer
                UINT32 clientFramesAvailable = bufferFrameCount - numFramesPadding;

                // Check that there's enough space in the audio client to take in all the data obtained from the loopback interface
                if (clientFramesAvailable < FramesAvailable)
                {
                    std::cout << "No space available in the render client to play back all the captured audio frames" << std::endl;
                }
                else
                {
                    // Grab all the available space in the shared buffer.
                    BYTE* pData = NULL;
                    float samplingRatio = (float)m_pOutputFormat->Format.nSamplesPerSec / (float)m_CaptureFormat.nSamplesPerSec;
                    hr = m_OutputRenderClient->GetBuffer((int)(FramesAvailable * samplingRatio)+1, &pData);
                    RETURN_IF_FAILED(hr);

                    // Resample the audio stream to the desired output format
                    UINT32 framesWritten = 0;
                    resampleAudioStream(Data, pData, FramesAvailable, clientFramesAvailable, framesWritten);
                    std::cout << "Resampled " << FramesAvailable << " frames into " << framesWritten << " frames. Requested frames " << (int)(FramesAvailable * samplingRatio) +1<< std::endl;

                    // Release the render client's buffer back
                    hr = m_OutputRenderClient->ReleaseBuffer(framesWritten, 0);
                    RETURN_IF_FAILED(hr);
                }

            }

            // Release the loopback capture's buffer back
            hr = m_AudioCaptureClient->ReleaseBuffer(FramesAvailable);
            RETURN_IF_FAILED(hr);

            // Increase the size of our 'data' chunk.  m_cbDataSize needs to be accurate
            m_cbDataSize += cbBytesToCapture;
        }

        if (FramesAvailable == 0)
        {
            std::cout << "There're no more frames to read" << std::endl;
            Sleep(1);
        }
    }

    return S_OK;
}