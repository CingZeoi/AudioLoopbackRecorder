/*
	* CLoopbackCapture��ʵ����Ƶ���ز����ܣ�֧�ֲ���ϵͳȫ����Ƶ��ָ�����̵���Ƶ�����WAV�ļ���
	* ��Ҫ���ܣ�
	*   - ��ʼ����Ƶ���񻷾����¼����������еȣ�
	*   - ������Ƶ�ӿڣ�ȫ�ֻ�����ض���
	*   - ������Ƶ��ʽ�ͻ�����
	*   - ����WAV�ļ���д����Ƶ����
	*   - ʹ�ö��̴߳�����Ƶ������ļ�д��
	*   - ��ȷֹͣ�����޸�WAV�ļ�ͷ
	*
	* ʵ��˼·��
	*   1. ��ʼ���׶Σ�������Ҫ���¼���������Media Foundation����ȡ��������
	*   2. ������Ƶ�ӿڣ�ͨ��ϵͳAPI��ȡ��Ƶ�ͻ��˽ӿ�
	*   3. ������Ƶ��ʽ������PCM��ʽ�����������ʡ�λ��������ȣ�
	*   4. �ļ�׼��������WAV�ļ���д���ʼ�ļ�ͷ
	*   5. ����׶Σ�������Ƶ�ͻ��ˣ����������̴߳�����Ƶ����
	*   6. ֹͣ�׶Σ�ֹͣ���񣬵ȴ�д���߳���ɣ��޸�WAV�ļ�ͷ
	*
	* ʹ�õ��ļ���/�⣺
	*   - Windows��ƵAPI��AudioClient.h, mmdeviceapi.h��
	*   - Media Foundation API��mfapi.h��
	*   - Windows Implementation Library��WIL��- ��COM����Դ����
	*   - C++��׼�̺߳�ͬ��ԭ��
	*
	* ע�⣺
	*   - ʹ��COM�������Ҫ��ȷ�������ü����ͽӿڲ�ѯ
	*   - ���̻߳�������Ҫ��������ͬ����ʹ�û���������������
	*   - WAV�ļ�ͷ��Ҫ�ڲ�����ɺ���������д����ȷ�����ݴ�С
	*/

#include <shlobj.h>      // Shell��ع���
#include <wchar.h>       // ���ַ�����
#include <iostream>      // ���������
#include <audioclientactivationparams.h>  // ��Ƶ�ͻ��˼������

#include "LoopbackCapture.h"  // ���ز���ͷ�ļ�

#define BITS_PER_BYTE 8  // ����ÿ�ֽ�λ��

	// ���캯������ʼ��ԭ�ӱ����ͳ�Ա����
CLoopbackCapture::CLoopbackCapture() :
	m_bIsCapturing(false),      // ��ʼ������״̬Ϊfalse
	m_writerThreadResult(S_OK)  // ��ʼ��д���߳̽��Ϊ�ɹ�
{
}

// �����豸״̬Ϊ�����������ʧ�ܣ�
HRESULT CLoopbackCapture::SetDeviceStateErrorIfFailed(HRESULT hr)
{
	if (FAILED(hr))  // ���HRESULT�Ƿ��ʾʧ��
	{
		m_DeviceState = DeviceState::Error;  // �����豸״̬Ϊ����
	}
	return hr;  // ����ԭʼHRESULT
}

// ��ʼ�����ز��񻷾�
HRESULT CLoopbackCapture::InitializeLoopbackCapture()
{
	// �������������¼��������첽֪ͨ��
	RETURN_IF_FAILED(m_SampleReadyEvent.create(wil::EventOptions::None));

	// ����Media Foundation��������ģʽ��
	RETURN_IF_FAILED(MFStartup(MF_VERSION, MFSTARTUP_LITE));

	// ��ȡ���������������첽����
	DWORD dwTaskID = 0;
	RETURN_IF_FAILED(MFLockSharedWorkQueue(L"Capture", 0, &dwTaskID, &m_dwQueueID));

	// �������������ص��Ķ���ID
	m_xSampleReady.SetQueueID(m_dwQueueID);

	// ������������¼�
	RETURN_IF_FAILED(m_hActivateCompleted.create(wil::EventOptions::None));

	// ��������ֹͣ�¼�
	RETURN_IF_FAILED(m_hCaptureStopped.create(wil::EventOptions::None));

	return S_OK;
}

// ����������������Դ
CLoopbackCapture::~CLoopbackCapture()
{
	// ����д���߳�
	if (m_WriterThread.joinable())  // ����߳��Ƿ������
	{
		if (m_bIsCapturing)  // ������ڲ�����
		{
			StopCaptureAsync();  // �첽ֹͣ����
		}
		else
		{
			m_WriterThread.join();  // �ȴ��߳̽���
		}
	}

	// ������������
	if (m_dwQueueID != 0)
	{
		MFUnlockWorkQueue(m_dwQueueID);
	}
}

// ����ָ�����̵���Ƶ�ӿ�
HRESULT CLoopbackCapture::ActivateAudioInterface(DWORD processId, bool includeProcessTree)
{
	return SetDeviceStateErrorIfFailed([&]() -> HRESULT
		{
			// ������Ƶ�ͻ��˼������
			AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = {};
			audioclientActivationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;  // ���̻���ģʽ
			audioclientActivationParams.ProcessLoopbackParams.ProcessLoopbackMode = includeProcessTree ?
				PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE : PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE;  // �������ų�������
			audioclientActivationParams.ProcessLoopbackParams.TargetProcessId = processId;  // Ŀ�����ID

			// �������Ա���
			PROPVARIANT activateParams = {};
			activateParams.vt = VT_BLOB;  // ����Ϊ�����ƴ����
			activateParams.blob.cbSize = sizeof(audioclientActivationParams);  // ���ݴ�С
			activateParams.blob.pBlobData = (BYTE*)&audioclientActivationParams;  // ����ָ��

			// �첽������Ƶ�ӿ�
			wil::com_ptr_nothrow<IActivateAudioInterfaceAsyncOperation> asyncOp;
			RETURN_IF_FAILED(ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &activateParams, this, &asyncOp));

			// �ȴ��������
			m_hActivateCompleted.wait();

			return m_activateResult;  // ���ؼ�����
		}());
}

// ��Ƶ�ӿڼ�����ɻص�
HRESULT CLoopbackCapture::ActivateCompleted(IActivateAudioInterfaceAsyncOperation* operation)
{
	// ��������
	m_activateResult = SetDeviceStateErrorIfFailed([&]()->HRESULT
		{
			HRESULT hrActivateResult = E_UNEXPECTED;
			wil::com_ptr_nothrow<IUnknown> punkAudioInterface;

			// ��ȡ������
			RETURN_IF_FAILED(operation->GetActivateResult(&hrActivateResult, &punkAudioInterface));
			RETURN_IF_FAILED(hrActivateResult);

			// ��ȡ��Ƶ�ͻ��˽ӿ�
			RETURN_IF_FAILED(punkAudioInterface.copy_to(&m_AudioClient));

			// ������Ƶ��ʽ��PCM��44.1kHz��16λ����������
			m_CaptureFormat.wFormatTag = WAVE_FORMAT_PCM;
			m_CaptureFormat.nChannels = 2;
			m_CaptureFormat.nSamplesPerSec = 44100;
			m_CaptureFormat.wBitsPerSample = 16;
			m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / BITS_PER_BYTE;
			m_CaptureFormat.nAvgBytesPerSec = m_CaptureFormat.nSamplesPerSec * m_CaptureFormat.nBlockAlign;

			// ��ʼ����Ƶ�ͻ���
			RETURN_IF_FAILED(m_AudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED,
				AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
				200000,  // ����������ʱ�䣨200���룩
				AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,  // �Զ�ת��PCM��ʽ
				&m_CaptureFormat,
				nullptr));

			// ��ȡ��������С
			RETURN_IF_FAILED(m_AudioClient->GetBufferSize(&m_BufferFrames));

			// ��ȡ��Ƶ����ͻ���
			RETURN_IF_FAILED(m_AudioClient->GetService(IID_PPV_ARGS(&m_AudioCaptureClient)));

			// �����첽�������
			RETURN_IF_FAILED(MFCreateAsyncResult(nullptr, &m_xSampleReady, nullptr, &m_SampleReadyAsyncResult));

			// �����¼����
			RETURN_IF_FAILED(m_AudioClient->SetEventHandle(m_SampleReadyEvent.get()));

			// ����WAV�ļ�
			RETURN_IF_FAILED(CreateWAVFile());

			// �����豸״̬Ϊ�ѳ�ʼ��
			m_DeviceState = DeviceState::Initialized;
			return S_OK;
		}());

	// ���ü�������¼�
	m_hActivateCompleted.SetEvent();
	return S_OK;
}

// ����WAV�ļ���д���ļ�ͷ
HRESULT CLoopbackCapture::CreateWAVFile()
{
	return SetDeviceStateErrorIfFailed([&]()->HRESULT
		{
			// �����ļ�
			m_hFile.reset(CreateFile(m_outputFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
			RETURN_LAST_ERROR_IF(!m_hFile);

			// д��RIFF��fmt��ͷ
			DWORD header[] = {
							FCC('RIFF'), 0, FCC('WAVE'), FCC('fmt '), sizeof(m_CaptureFormat)
			};
			DWORD dwBytesWritten = 0;
			RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), header, sizeof(header), &dwBytesWritten, NULL));
			m_cbHeaderSize += dwBytesWritten;

			// д����Ƶ��ʽ��Ϣ
			WI_ASSERT(m_CaptureFormat.cbSize == 0);
			RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_CaptureFormat, sizeof(m_CaptureFormat), &dwBytesWritten, NULL));
			m_cbHeaderSize += dwBytesWritten;

			// д�����ݿ�ͷ
			DWORD data[] = { FCC('data'), 0 };
			RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), data, sizeof(data), &dwBytesWritten, NULL));
			m_cbHeaderSize += dwBytesWritten;

			return S_OK;
		}());
}

// �޸�WAV�ļ�ͷ��д����ȷ�����ݴ�С��
HRESULT CLoopbackCapture::FixWAVHeader()
{
	// ��λ�����ݴ�С�ֶβ�д��ʵ�����ݴ�С
	DWORD dwPtr = SetFilePointer(m_hFile.get(), m_cbHeaderSize - sizeof(DWORD), NULL, FILE_BEGIN);
	RETURN_LAST_ERROR_IF(INVALID_SET_FILE_POINTER == dwPtr);
	DWORD dwBytesWritten = 0;
	RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_cbDataSize, sizeof(DWORD), &dwBytesWritten, NULL));

	// ��λ���ļ��ܴ�С�ֶβ�д����ȷֵ
	RETURN_LAST_ERROR_IF(INVALID_SET_FILE_POINTER == SetFilePointer(m_hFile.get(), sizeof(DWORD), NULL, FILE_BEGIN));
	DWORD cbTotalSize = m_cbDataSize + m_cbHeaderSize - 8;
	RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &cbTotalSize, sizeof(DWORD), &dwBytesWritten, NULL));

	// ˢ���ļ�������
	RETURN_IF_WIN32_BOOL_FALSE(FlushFileBuffers(m_hFile.get()));
	return S_OK;
}

// ��ʼ����ָ�����̵���Ƶ
HRESULT CLoopbackCapture::StartCaptureAsync(DWORD processId, bool includeProcessTree, PCWSTR outputFileName)
{
	m_outputFileName = outputFileName;
	// ʹ���������˳�ȷ���ļ���������
	auto resetOutputFileName = wil::scope_exit([&] { m_outputFileName = nullptr; });

	// ��ʼ�����񻷾�
	RETURN_IF_FAILED(InitializeLoopbackCapture());

	// ������Ƶ�ӿ�
	RETURN_IF_FAILED(ActivateAudioInterface(processId, includeProcessTree));

	// ����豸�ѳ�ʼ������ʼ����
	if (m_DeviceState == DeviceState::Initialized)
	{
		m_DeviceState = DeviceState::Starting;
		// ����ʼ����������빤������
		return MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStartCapture, nullptr);
	}
	return S_OK;
}

// ��ʼȫ����Ƶ����
HRESULT CLoopbackCapture::StartGlobalCaptureAsync(PCWSTR outputFileName)
{
	m_outputFileName = outputFileName;
	// ʹ���������˳�ȷ���ļ���������
	auto resetOutputFileName = wil::scope_exit([&] { m_outputFileName = nullptr; });

	// ��ʼ�����񻷾�
	RETURN_IF_FAILED(InitializeLoopbackCapture());

	// ����ȫ����Ƶ�ӿ�
	RETURN_IF_FAILED(ActivateAudioInterfaceGlobal());

	// ����豸�ѳ�ʼ������ʼ����
	if (m_DeviceState == DeviceState::Initialized)
	{
		m_DeviceState = DeviceState::Starting;
		// ����ʼ����������빤������
		return MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStartCapture, nullptr);
	}
	return S_OK;
}

// ����ȫ����Ƶ�ӿ�
HRESULT CLoopbackCapture::ActivateAudioInterfaceGlobal()
{
	return SetDeviceStateErrorIfFailed([&]() -> HRESULT
		{
			wil::com_ptr_nothrow<IMMDeviceEnumerator> enumerator;
			wil::com_ptr_nothrow<IMMDevice> device;

			// �����豸ö����
			RETURN_IF_FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&enumerator)));

			// ��ȡĬ����Ƶ��Ⱦ�˵�
			RETURN_IF_FAILED(enumerator->GetDefaultAudioEndpoint(eRender, eConsole, &device));

			// ������Ƶ�ͻ��˽ӿ�
			RETURN_IF_FAILED(device->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void**)&m_AudioClient));

			// ������Ƶ��ʽ��PCM��44.1kHz��16λ����������
			m_CaptureFormat.wFormatTag = WAVE_FORMAT_PCM;
			m_CaptureFormat.nChannels = 2;
			m_CaptureFormat.nSamplesPerSec = 44100;
			m_CaptureFormat.wBitsPerSample = 16;
			m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / BITS_PER_BYTE;
			m_CaptureFormat.nAvgBytesPerSec = m_CaptureFormat.nSamplesPerSec * m_CaptureFormat.nBlockAlign;
			m_CaptureFormat.cbSize = 0;

			// ��ʼ����Ƶ�ͻ��ˣ�����ģʽ��
			RETURN_IF_FAILED(m_AudioClient->Initialize(
				AUDCLNT_SHAREMODE_SHARED,
				AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
				200000,  // ����������ʱ�䣨200���룩
				0,
				&m_CaptureFormat,
				nullptr));

			// ��ȡ��������С
			RETURN_IF_FAILED(m_AudioClient->GetBufferSize(&m_BufferFrames));

			// ��ȡ��Ƶ����ͻ���
			RETURN_IF_FAILED(m_AudioClient->GetService(IID_PPV_ARGS(&m_AudioCaptureClient)));

			// �����첽�������
			RETURN_IF_FAILED(MFCreateAsyncResult(nullptr, &m_xSampleReady, nullptr, &m_SampleReadyAsyncResult));

			// �����¼����
			RETURN_IF_FAILED(m_AudioClient->SetEventHandle(m_SampleReadyEvent.get()));

			// ����WAV�ļ�
			RETURN_IF_FAILED(CreateWAVFile());

			// �����豸״̬Ϊ�ѳ�ʼ��
			m_DeviceState = DeviceState::Initialized;
			return S_OK;
		}());
}

// ��ʼ����ص�
HRESULT CLoopbackCapture::OnStartCapture(IMFAsyncResult* pResult)
{
	return SetDeviceStateErrorIfFailed([&]()->HRESULT
		{
			// ������Ƶ�ͻ���
			RETURN_IF_FAILED(m_AudioClient->Start());

			// �����豸״̬Ϊ������
			m_DeviceState = DeviceState::Capturing;

			// ����д���߳�
			m_bIsCapturing = true;
			m_writerThreadResult = S_OK;
			m_WriterThread = std::thread(&CLoopbackCapture::WriterThreadProc, this);

			// �������������������ȴ�����
			MFPutWaitingWorkItem(m_SampleReadyEvent.get(), 0, m_SampleReadyAsyncResult.get(), &m_SampleReadyKey);
			return S_OK;
		}());
}

// �첽ֹͣ����
HRESULT CLoopbackCapture::StopCaptureAsync()
{
	// ����豸״̬�Ƿ���Ч
	RETURN_HR_IF(E_NOT_VALID_STATE, (m_DeviceState != DeviceState::Capturing) && (m_DeviceState != DeviceState::Error));

	// ����Ѿ���ֹͣ����ֹͣ��ֱ�ӷ���
	if (m_DeviceState == DeviceState::Stopping || m_DeviceState == DeviceState::Stopped)
	{
		return S_OK;
	}

	// �����豸״̬Ϊֹͣ��
	m_DeviceState = DeviceState::Stopping;

	// ��ֹͣ����������빤������
	RETURN_IF_FAILED(MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStopCapture, nullptr));

	// �ȴ�������ȫֹͣ
	m_hCaptureStopped.wait();

	// �ȴ�д���߳̽���
	if (m_WriterThread.joinable())
	{
		m_WriterThread.join();
	}

	// �����豸״̬Ϊ��ֹͣ
	m_DeviceState = DeviceState::Stopped;

	// ����д���̵߳Ľ��
	return m_writerThreadResult;
}

// ֹͣ����ص�
HRESULT CLoopbackCapture::OnStopCapture(IMFAsyncResult* pResult)
{
	// ȡ����������������
	if (0 != m_SampleReadyKey)
	{
		MFCancelWorkItem(m_SampleReadyKey);
		m_SampleReadyKey = 0;
	}

	// ֹͣ��Ƶ�ͻ���
	m_AudioClient->Stop();

	// �����첽�������
	m_SampleReadyAsyncResult.reset();

	// ���²���״̬
	m_bIsCapturing = false;

	// ֪ͨд���߳�
	m_QueueCV.notify_one();

	// ���ò���ֹͣ�¼�
	m_hCaptureStopped.SetEvent();

	return S_OK;
}

// ���������ص�
HRESULT CLoopbackCapture::OnSampleReady(IMFAsyncResult* pResult)
{
	// ������Ƶ��������
	if (SUCCEEDED(OnAudioSampleRequested()))
	{
		// ������ڲ����У������ȴ���һ������
		if (m_DeviceState == DeviceState::Capturing)
		{
			return MFPutWaitingWorkItem(m_SampleReadyEvent.get(), 0, m_SampleReadyAsyncResult.get(), &m_SampleReadyKey);
		}
	}
	else
	{
		// �������ʧ�ܣ������豸״̬Ϊ����
		m_DeviceState = DeviceState::Error;
	}
	return S_OK;
}

// ������Ƶ��������
HRESULT CLoopbackCapture::OnAudioSampleRequested()
{
	UINT32 FramesAvailable = 0;
	BYTE* Data = nullptr;
	DWORD dwCaptureFlags;
	UINT64 u64DevicePosition = 0;
	UINT64 u64QPCPosition = 0;

	// ��ȡ�ٽ�����
	auto lock = m_CritSec.lock();

	// ����豸״̬
	if (m_DeviceState == DeviceState::Stopping || m_DeviceState == DeviceState::Stopped)
	{
		return S_OK;
	}

	// �������п��õ���Ƶ���ݰ�
	while (SUCCEEDED(m_AudioCaptureClient->GetNextPacketSize(&FramesAvailable)) && FramesAvailable > 0)
	{
		// ������Ҫ������ֽ���
		UINT32 cbBytesToCapture = FramesAvailable * m_CaptureFormat.nBlockAlign;

		// ��ȡ��Ƶ������
		RETURN_IF_FAILED(m_AudioCaptureClient->GetBuffer(&Data, &FramesAvailable, &dwCaptureFlags, &u64DevicePosition, &u64QPCPosition));

		try
		{
			// ������Ƶ���ݵ�����
			std::vector<BYTE> audioChunk(Data, Data + cbBytesToCapture);

			// ����Ƶ���ݷ������
			{
				std::lock_guard<std::mutex> queueLock(m_QueueMutex);
				m_AudioQueue.push(std::move(audioChunk));
			}

			// ֪ͨд���߳���������
			m_QueueCV.notify_one();
		}
		catch (const std::bad_alloc&)
		{
			// �ڴ����ʧ�ܴ���
			m_writerThreadResult = E_OUTOFMEMORY;
			StopCaptureAsync();
			break;
		}

		// �ͷ���Ƶ������
		m_AudioCaptureClient->ReleaseBuffer(FramesAvailable);
	}
	return S_OK;
}

// д���̴߳�����
void CLoopbackCapture::WriterThreadProc()
{
	// ѭ��������Ƶ���ݣ�ֱ������ֹͣ�Ҷ���Ϊ��
	while (m_bIsCapturing || !m_AudioQueue.empty())
	{
		std::vector<BYTE> audioData;

		// �Ӷ����л�ȡ��Ƶ����
		{
			std::unique_lock<std::mutex> lock(m_QueueMutex);

			// �ȴ������������ݻ򲶻�ֹͣ
			m_QueueCV.wait(lock, [this] {
				return !m_AudioQueue.empty() || !m_bIsCapturing;
				});

			// �������Ϊ�գ������ȴ�
			if (m_AudioQueue.empty())
			{
				continue;
			}

			// ��ȡ����ǰ�������
			audioData = std::move(m_AudioQueue.front());
			m_AudioQueue.pop();
		}

		// ����Ƶ����д���ļ�
		if (!audioData.empty())
		{
			DWORD dwBytesWritten = 0;
			if (!WriteFile(m_hFile.get(), audioData.data(), static_cast<DWORD>(audioData.size()), &dwBytesWritten, NULL))
			{
				// д��ʧ�ܴ���
				m_writerThreadResult = HRESULT_FROM_WIN32(GetLastError());
				m_bIsCapturing = false;
				continue;
			}

			// ������д�����ݴ�С
			m_cbDataSize += dwBytesWritten;
		}
	}

	// ������ɺ��޸�WAV�ļ�ͷ
	if (SUCCEEDED(m_writerThreadResult))
	{
		HRESULT hr = FixWAVHeader();
		if (FAILED(hr))
		{
			m_writerThreadResult = hr;
		}
	}
}