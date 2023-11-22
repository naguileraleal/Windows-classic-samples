#include "LoopbackCaptureBase.h"

#include <iostream>

/**
* Initializes a Media Foundation audio resampler
* Taken from https://sourceforge.net/p/playpcmwin/wiki/HowToUseResamplerMFT/
* TODO: Handle errors
*/
HRESULT LoopbackCaptureBase::initializeMFTResampler(WAVEFORMATEX* inputFmt, WAVEFORMATEXTENSIBLE* outputFmtex)
{
	HRESULT hr = S_OK;
	CComPtr<IUnknown> spTransformUnk;
	CComPtr<IWMResamplerProps> spResamplerProps;
	IMFMediaType *pInputType = NULL, *pOutputType = NULL;
	WAVEFORMATEXTENSIBLE* fmtex;

	hr = MFStartup(MF_VERSION, MFSTARTUP_NOSOCKET);
	EXIT_ON_ERROR(hr)


	hr = CoCreateInstance(CLSID_CResamplerMediaObject, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (void**)&spTransformUnk);

	hr = spTransformUnk->QueryInterface(IID_PPV_ARGS(&m_ResamplerTransform));
	EXIT_ON_ERROR(hr)

		hr = spTransformUnk->QueryInterface(IID_PPV_ARGS(&spResamplerProps));
	hr = spResamplerProps->SetHalfFilterLength(60); //< best conversion quality. Use 1 for lowest quality (and lowest latency)

	// Specify input/output formats
	// set input format
	MFCreateMediaType(&pInputType);
	hr = pInputType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Audio);
	hr = pInputType->SetGUID(MF_MT_SUBTYPE, (inputFmt->wFormatTag == WAVE_FORMAT_PCM) ? MFAudioFormat_PCM : MFAudioFormat_Float);
	hr = pInputType->SetUINT32(MF_MT_AUDIO_NUM_CHANNELS,         inputFmt->nChannels);
	hr = pInputType->SetUINT32(MF_MT_AUDIO_SAMPLES_PER_SECOND,   inputFmt->nSamplesPerSec);
	hr = pInputType->SetUINT32(MF_MT_AUDIO_BLOCK_ALIGNMENT,      inputFmt->nBlockAlign);
	hr = pInputType->SetUINT32(MF_MT_AUDIO_AVG_BYTES_PER_SECOND, inputFmt->nAvgBytesPerSec);
	hr = pInputType->SetUINT32(MF_MT_AUDIO_BITS_PER_SAMPLE,      inputFmt->wBitsPerSample);
	hr = pInputType->SetUINT32(MF_MT_ALL_SAMPLES_INDEPENDENT,    TRUE);
	//if (0 != inputFmtex->dwChannelMask) {
	//	hr = pInputType->SetUINT32(MF_MT_AUDIO_CHANNEL_MASK, inputFmtex->dwChannelMask);
	//}
	//if (fmt.wBitsPerSample != inputFmtex->Samples.wValidBitsPerSample) {
	//	hr = pInputType->SetUINT32(MF_MT_AUDIO_VALID_BITS_PER_SAMPLE, inputFmtex->Samples.wValidBitsPerSample);
	//}
	hr = m_ResamplerTransform->SetInputType(0, pInputType, 0);
	// Set output format
	MFCreateMediaType(&pOutputType);
	hr = pOutputType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Audio);
	hr = pOutputType->SetGUID(MF_MT_SUBTYPE, (outputFmtex->Format.wFormatTag == WAVE_FORMAT_PCM) ? MFAudioFormat_PCM : MFAudioFormat_Float);
	hr = pOutputType->SetUINT32(MF_MT_AUDIO_NUM_CHANNELS, outputFmtex->Format.nChannels);
	hr = pOutputType->SetUINT32(MF_MT_AUDIO_SAMPLES_PER_SECOND, outputFmtex->Format.nSamplesPerSec);
	hr = pOutputType->SetUINT32(MF_MT_AUDIO_BLOCK_ALIGNMENT, outputFmtex->Format.nBlockAlign);
	hr = pOutputType->SetUINT32(MF_MT_AUDIO_AVG_BYTES_PER_SECOND, outputFmtex->Format.nAvgBytesPerSec);
	hr = pOutputType->SetUINT32(MF_MT_AUDIO_BITS_PER_SAMPLE, outputFmtex->Format.wBitsPerSample);
	hr = pOutputType->SetUINT32(MF_MT_ALL_SAMPLES_INDEPENDENT, TRUE);
	if (0 != outputFmtex->dwChannelMask) {
		hr = pOutputType->SetUINT32(MF_MT_AUDIO_CHANNEL_MASK, outputFmtex->dwChannelMask);
	}
	if (outputFmtex->Format.wBitsPerSample != outputFmtex->Samples.wValidBitsPerSample) {
		hr = pOutputType->SetUINT32(MF_MT_AUDIO_VALID_BITS_PER_SAMPLE, outputFmtex->Samples.wValidBitsPerSample);
	}
	hr = m_ResamplerTransform->SetOutputType(0, pOutputType, 0);

Exit:
	return hr;
}

void LoopbackCaptureBase::resampleAudioStream(BYTE* src, BYTE* dst, UINT32 framesAvailable, UINT32 clientFramesAvailable, UINT32& framesWritten)
{
	if (m_ResamplerTransform != nullptr)
	{
		BYTE  *data = src; //< input PCM data 
		DWORD bytes = framesAvailable * m_CaptureFormat.nBlockAlign; //< bytes need to be smaller than approx. 1Mbytes
		HRESULT hr = S_OK;

		MFT_OUTPUT_STREAM_INFO outputStreamInfo;
		m_ResamplerTransform->GetOutputStreamInfo(0, &outputStreamInfo);

		CComPtr<IMFMediaBuffer> pBuffer = NULL;
		hr = MFCreateMemoryBuffer(bytes , &pBuffer);

		BYTE  *pByteBufferTo = NULL;
		hr = pBuffer->Lock(&pByteBufferTo, NULL, NULL);
		memcpy(pByteBufferTo, data, bytes);
		pBuffer->Unlock();
		pByteBufferTo = NULL;

		hr = pBuffer->SetCurrentLength(bytes);

		IMFSample *pSample = NULL;
		hr = MFCreateSample(&pSample);
		hr = pSample->AddBuffer(pBuffer);

		hr = m_ResamplerTransform->ProcessInput(0, pSample, 0);

		// Allocate a buffer to receive output from media transform
		MFT_OUTPUT_DATA_BUFFER outputDataBuffer;
		outputDataBuffer.dwStreamID = 0;
		outputDataBuffer.dwStatus = 0;
		outputDataBuffer.pEvents = NULL;

		CComPtr<IMFSample> pOutSample = NULL;
		hr = MFCreateSample(&pOutSample);
		DWORD OutputMediaBufferCapacity = m_pOutputFormat->Format.nAvgBytesPerSec;
		CComPtr<IMFMediaBuffer> outSampleBuffer = NULL;
		MFCreateMemoryBuffer(OutputMediaBufferCapacity, &outSampleBuffer);
		outSampleBuffer->SetCurrentLength(0);
		pOutSample->AddBuffer(outSampleBuffer);
		outputDataBuffer.pSample = pOutSample;
		DWORD dwStatus;
		hr = m_ResamplerTransform->ProcessOutput(0, 1, &outputDataBuffer, &dwStatus);
		//if (hr == MF_E_TRANSFORM_NEED_MORE_INPUT) {
		//	// conversion end
		//	break;
		//}

		DWORD cbBytes = 0;
		BYTE  *pByteBuffer = NULL;
		hr = outSampleBuffer->GetCurrentLength(&cbBytes);
		outSampleBuffer->Lock(&pByteBuffer, NULL, NULL);
		memcpy(dst, pByteBuffer, cbBytes);
		outSampleBuffer->Unlock();

		framesWritten = cbBytes / m_pOutputFormat->Format.nBlockAlign;
	}
	// If MFT is not available, use linear interpolation
	else if (m_pOutputFormat != nullptr)
	{
		// Resample the audio stream to the desired output format
		if (m_pOutputFormat->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT)
		{
			if (m_pOutputFormat->Format.wBitsPerSample == 32)
			{

				// TODO: DEBUGGING! DELETE TIMING CODE!
				LARGE_INTEGER frequency, startTime, endTime;
				QueryPerformanceFrequency(&frequency);
				QueryPerformanceCounter(&startTime);
				// Convert from 16-bit PCM to 32-bit float. No sample rate change!
				//for (int i = 0; i < framesAvailable; i++)
				//{
				//	int16_t* srcPtr = (int16_t*)(src + i * m_CaptureFormat.nBlockAlign);
				//	float* dstPtr = (float*)(dst + i * m_pOutputFormat->Format.nBlockAlign);

				//	for (auto channel = 0; channel < m_CaptureFormat.nChannels; channel++)
				//	{
				//		*dstPtr = (float)(*srcPtr / 32768.0f);
				//		dstPtr++;
				//		srcPtr++;
				//	}
				//}
				// Perform linear interpolation if needed
				float sampleRatio = (float)m_pOutputFormat->Format.nSamplesPerSec / (float)m_CaptureFormat.nSamplesPerSec; // Ratio between the two sample rates
				auto srcSamples = framesAvailable * m_CaptureFormat.nChannels; // Number of bytes in the input buffer
				int dstFrames = (int)(framesAvailable * sampleRatio); // Number of frames in the output buffer
				framesWritten = dstFrames;
				int dstSamples = dstFrames * m_pOutputFormat->Format.nChannels; // Number of bytes in the output buffer

				for (auto i = 0; i < dstSamples; i++)
				{
					if (i >= clientFramesAvailable)
					{
						// Avoid overwriting buffer
						break;
					}
					float realPos = i / sampleRatio;
					int iLow = (int)realPos;
					int iHigh = iLow + 1;
					float remainder = realPos - (float)iLow;

					float lowval = 0;
					float highval = 0;
					if ((iLow >= 0) && (iLow < srcSamples))
					{
						lowval = ((int16_t*)src)[iLow] / 32768.0f;
					}
					if ((iHigh >= 0) && (iHigh < srcSamples))
					{
						highval = ((int16_t*)src)[iHigh] / 32768.0f;
					}

					((float*)dst)[i] = (highval * remainder) + (lowval * (1 - remainder));
				}


				QueryPerformanceCounter(&endTime);
				LONGLONG elapsedTime = ((endTime.QuadPart - startTime.QuadPart) * 1000000) / frequency.QuadPart;
				std::cout << "Execution took " << elapsedTime << "us" << std::endl;
				// END TIMING BLOCK


			}
			else if (m_pOutputFormat->Format.wBitsPerSample == 64)
			{
				// Convert from 16-bit PCM to 64-bit float
				for (int i = 0; i < framesAvailable; i++)
				{
					int16_t* srcPtr = (int16_t*)(src + i * m_CaptureFormat.nBlockAlign);
					double* dstPtr = (double*)(dst + i * m_pOutputFormat->Format.nBlockAlign);
					for (auto channel = 0; channel < m_CaptureFormat.nChannels; channel++)
					{
						*dstPtr = (double)(*srcPtr / 32768.0f);
						dstPtr++;
						srcPtr++;
					}
				}
			}
		}
	}
}