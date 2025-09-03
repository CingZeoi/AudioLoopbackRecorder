/*
	* CLoopbackCapture类实现音频环回捕获功能，支持捕获系统全局音频或指定进程的音频输出到WAV文件。
	* 主要功能：
	*   - 初始化音频捕获环境（事件、工作队列等）
	*   - 激活音频接口（全局或进程特定）
	*   - 配置音频格式和缓冲区
	*   - 创建WAV文件并写入音频数据
	*   - 使用多线程处理音频捕获和文件写入
	*   - 正确停止捕获并修复WAV文件头
	*
	* 实现思路：
	*   1. 初始化阶段：创建必要的事件对象，启动Media Foundation，获取工作队列
	*   2. 激活音频接口：通过系统API获取音频客户端接口
	*   3. 配置音频格式：设置PCM格式参数（采样率、位深、声道数等）
	*   4. 文件准备：创建WAV文件并写入初始文件头
	*   5. 捕获阶段：启动音频客户端，创建工作线程处理音频数据
	*   6. 停止阶段：停止捕获，等待写入线程完成，修复WAV文件头
	*
	* 使用到的技术/库：
	*   - Windows音频API（AudioClient.h, mmdeviceapi.h）
	*   - Media Foundation API（mfapi.h）
	*   - Windows Implementation Library（WIL）- 简化COM和资源管理
	*   - C++标准线程和同步原语
	*
	* 注意：
	*   - 使用COM组件，需要正确管理引用计数和接口查询
	*   - 多线程环境下需要谨慎处理同步，使用互斥锁和条件变量
	*   - WAV文件头需要在捕获完成后修正，以写入正确的数据大小
	*/

#include <shlobj.h>      // Shell相关功能
#include <wchar.h>       // 宽字符处理
#include <iostream>      // 输入输出流
#include <audioclientactivationparams.h>  // 音频客户端激活参数

#include "LoopbackCapture.h"  // 环回捕获头文件

#define BITS_PER_BYTE 8  // 定义每字节位数

	// 构造函数：初始化原子变量和成员变量
CLoopbackCapture::CLoopbackCapture() :
	m_bIsCapturing(false),      // 初始化捕获状态为false
	m_writerThreadResult(S_OK)  // 初始化写入线程结果为成功
{
}

// 设置设备状态为错误（如果操作失败）
HRESULT CLoopbackCapture::SetDeviceStateErrorIfFailed(HRESULT hr)
{
	if (FAILED(hr))  // 检查HRESULT是否表示失败
	{
		m_DeviceState = DeviceState::Error;  // 设置设备状态为错误
	}
	return hr;  // 返回原始HRESULT
}

// 初始化环回捕获环境
HRESULT CLoopbackCapture::InitializeLoopbackCapture()
{
	// 创建样本就绪事件（用于异步通知）
	RETURN_IF_FAILED(m_SampleReadyEvent.create(wil::EventOptions::None));

	// 启动Media Foundation（轻量级模式）
	RETURN_IF_FAILED(MFStartup(MF_VERSION, MFSTARTUP_LITE));

	// 获取共享工作队列用于异步操作
	DWORD dwTaskID = 0;
	RETURN_IF_FAILED(MFLockSharedWorkQueue(L"Capture", 0, &dwTaskID, &m_dwQueueID));

	// 设置样本就绪回调的队列ID
	m_xSampleReady.SetQueueID(m_dwQueueID);

	// 创建激活完成事件
	RETURN_IF_FAILED(m_hActivateCompleted.create(wil::EventOptions::None));

	// 创建捕获停止事件
	RETURN_IF_FAILED(m_hCaptureStopped.create(wil::EventOptions::None));

	return S_OK;
}

// 析构函数：清理资源
CLoopbackCapture::~CLoopbackCapture()
{
	// 处理写入线程
	if (m_WriterThread.joinable())  // 检查线程是否可连接
	{
		if (m_bIsCapturing)  // 如果仍在捕获中
		{
			StopCaptureAsync();  // 异步停止捕获
		}
		else
		{
			m_WriterThread.join();  // 等待线程结束
		}
	}

	// 解锁工作队列
	if (m_dwQueueID != 0)
	{
		MFUnlockWorkQueue(m_dwQueueID);
	}
}

// 激活指定进程的音频接口
HRESULT CLoopbackCapture::ActivateAudioInterface(DWORD processId, bool includeProcessTree)
{
	return SetDeviceStateErrorIfFailed([&]() -> HRESULT
		{
			// 设置音频客户端激活参数
			AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = {};
			audioclientActivationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;  // 进程环回模式
			audioclientActivationParams.ProcessLoopbackParams.ProcessLoopbackMode = includeProcessTree ?
				PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE : PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE;  // 包含或排除进程树
			audioclientActivationParams.ProcessLoopbackParams.TargetProcessId = processId;  // 目标进程ID

			// 设置属性变量
			PROPVARIANT activateParams = {};
			activateParams.vt = VT_BLOB;  // 类型为二进制大对象
			activateParams.blob.cbSize = sizeof(audioclientActivationParams);  // 数据大小
			activateParams.blob.pBlobData = (BYTE*)&audioclientActivationParams;  // 数据指针

			// 异步激活音频接口
			wil::com_ptr_nothrow<IActivateAudioInterfaceAsyncOperation> asyncOp;
			RETURN_IF_FAILED(ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &activateParams, this, &asyncOp));

			// 等待激活完成
			m_hActivateCompleted.wait();

			return m_activateResult;  // 返回激活结果
		}());
}

// 音频接口激活完成回调
HRESULT CLoopbackCapture::ActivateCompleted(IActivateAudioInterfaceAsyncOperation* operation)
{
	// 处理激活结果
	m_activateResult = SetDeviceStateErrorIfFailed([&]()->HRESULT
		{
			HRESULT hrActivateResult = E_UNEXPECTED;
			wil::com_ptr_nothrow<IUnknown> punkAudioInterface;

			// 获取激活结果
			RETURN_IF_FAILED(operation->GetActivateResult(&hrActivateResult, &punkAudioInterface));
			RETURN_IF_FAILED(hrActivateResult);

			// 获取音频客户端接口
			RETURN_IF_FAILED(punkAudioInterface.copy_to(&m_AudioClient));

			// 配置音频格式（PCM，44.1kHz，16位，立体声）
			m_CaptureFormat.wFormatTag = WAVE_FORMAT_PCM;
			m_CaptureFormat.nChannels = 2;
			m_CaptureFormat.nSamplesPerSec = 44100;
			m_CaptureFormat.wBitsPerSample = 16;
			m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / BITS_PER_BYTE;
			m_CaptureFormat.nAvgBytesPerSec = m_CaptureFormat.nSamplesPerSec * m_CaptureFormat.nBlockAlign;

			// 初始化音频客户端
			RETURN_IF_FAILED(m_AudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED,
				AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
				200000,  // 缓冲区持续时间（200毫秒）
				AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,  // 自动转换PCM格式
				&m_CaptureFormat,
				nullptr));

			// 获取缓冲区大小
			RETURN_IF_FAILED(m_AudioClient->GetBufferSize(&m_BufferFrames));

			// 获取音频捕获客户端
			RETURN_IF_FAILED(m_AudioClient->GetService(IID_PPV_ARGS(&m_AudioCaptureClient)));

			// 创建异步结果对象
			RETURN_IF_FAILED(MFCreateAsyncResult(nullptr, &m_xSampleReady, nullptr, &m_SampleReadyAsyncResult));

			// 设置事件句柄
			RETURN_IF_FAILED(m_AudioClient->SetEventHandle(m_SampleReadyEvent.get()));

			// 创建WAV文件
			RETURN_IF_FAILED(CreateWAVFile());

			// 更新设备状态为已初始化
			m_DeviceState = DeviceState::Initialized;
			return S_OK;
		}());

	// 设置激活完成事件
	m_hActivateCompleted.SetEvent();
	return S_OK;
}

// 创建WAV文件并写入文件头
HRESULT CLoopbackCapture::CreateWAVFile()
{
	return SetDeviceStateErrorIfFailed([&]()->HRESULT
		{
			// 创建文件
			m_hFile.reset(CreateFile(m_outputFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL));
			RETURN_LAST_ERROR_IF(!m_hFile);

			// 写入RIFF和fmt块头
			DWORD header[] = {
							FCC('RIFF'), 0, FCC('WAVE'), FCC('fmt '), sizeof(m_CaptureFormat)
			};
			DWORD dwBytesWritten = 0;
			RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), header, sizeof(header), &dwBytesWritten, NULL));
			m_cbHeaderSize += dwBytesWritten;

			// 写入音频格式信息
			WI_ASSERT(m_CaptureFormat.cbSize == 0);
			RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_CaptureFormat, sizeof(m_CaptureFormat), &dwBytesWritten, NULL));
			m_cbHeaderSize += dwBytesWritten;

			// 写入数据块头
			DWORD data[] = { FCC('data'), 0 };
			RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), data, sizeof(data), &dwBytesWritten, NULL));
			m_cbHeaderSize += dwBytesWritten;

			return S_OK;
		}());
}

// 修复WAV文件头（写入正确的数据大小）
HRESULT CLoopbackCapture::FixWAVHeader()
{
	// 定位到数据大小字段并写入实际数据大小
	DWORD dwPtr = SetFilePointer(m_hFile.get(), m_cbHeaderSize - sizeof(DWORD), NULL, FILE_BEGIN);
	RETURN_LAST_ERROR_IF(INVALID_SET_FILE_POINTER == dwPtr);
	DWORD dwBytesWritten = 0;
	RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &m_cbDataSize, sizeof(DWORD), &dwBytesWritten, NULL));

	// 定位到文件总大小字段并写入正确值
	RETURN_LAST_ERROR_IF(INVALID_SET_FILE_POINTER == SetFilePointer(m_hFile.get(), sizeof(DWORD), NULL, FILE_BEGIN));
	DWORD cbTotalSize = m_cbDataSize + m_cbHeaderSize - 8;
	RETURN_IF_WIN32_BOOL_FALSE(WriteFile(m_hFile.get(), &cbTotalSize, sizeof(DWORD), &dwBytesWritten, NULL));

	// 刷新文件缓冲区
	RETURN_IF_WIN32_BOOL_FALSE(FlushFileBuffers(m_hFile.get()));
	return S_OK;
}

// 开始捕获指定进程的音频
HRESULT CLoopbackCapture::StartCaptureAsync(DWORD processId, bool includeProcessTree, PCWSTR outputFileName)
{
	m_outputFileName = outputFileName;
	// 使用作用域退出确保文件名被重置
	auto resetOutputFileName = wil::scope_exit([&] { m_outputFileName = nullptr; });

	// 初始化捕获环境
	RETURN_IF_FAILED(InitializeLoopbackCapture());

	// 激活音频接口
	RETURN_IF_FAILED(ActivateAudioInterface(processId, includeProcessTree));

	// 如果设备已初始化，开始捕获
	if (m_DeviceState == DeviceState::Initialized)
	{
		m_DeviceState = DeviceState::Starting;
		// 将开始捕获工作项放入工作队列
		return MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStartCapture, nullptr);
	}
	return S_OK;
}

// 开始全局音频捕获
HRESULT CLoopbackCapture::StartGlobalCaptureAsync(PCWSTR outputFileName)
{
	m_outputFileName = outputFileName;
	// 使用作用域退出确保文件名被重置
	auto resetOutputFileName = wil::scope_exit([&] { m_outputFileName = nullptr; });

	// 初始化捕获环境
	RETURN_IF_FAILED(InitializeLoopbackCapture());

	// 激活全局音频接口
	RETURN_IF_FAILED(ActivateAudioInterfaceGlobal());

	// 如果设备已初始化，开始捕获
	if (m_DeviceState == DeviceState::Initialized)
	{
		m_DeviceState = DeviceState::Starting;
		// 将开始捕获工作项放入工作队列
		return MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStartCapture, nullptr);
	}
	return S_OK;
}

// 激活全局音频接口
HRESULT CLoopbackCapture::ActivateAudioInterfaceGlobal()
{
	return SetDeviceStateErrorIfFailed([&]() -> HRESULT
		{
			wil::com_ptr_nothrow<IMMDeviceEnumerator> enumerator;
			wil::com_ptr_nothrow<IMMDevice> device;

			// 创建设备枚举器
			RETURN_IF_FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&enumerator)));

			// 获取默认音频渲染端点
			RETURN_IF_FAILED(enumerator->GetDefaultAudioEndpoint(eRender, eConsole, &device));

			// 激活音频客户端接口
			RETURN_IF_FAILED(device->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void**)&m_AudioClient));

			// 配置音频格式（PCM，44.1kHz，16位，立体声）
			m_CaptureFormat.wFormatTag = WAVE_FORMAT_PCM;
			m_CaptureFormat.nChannels = 2;
			m_CaptureFormat.nSamplesPerSec = 44100;
			m_CaptureFormat.wBitsPerSample = 16;
			m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / BITS_PER_BYTE;
			m_CaptureFormat.nAvgBytesPerSec = m_CaptureFormat.nSamplesPerSec * m_CaptureFormat.nBlockAlign;
			m_CaptureFormat.cbSize = 0;

			// 初始化音频客户端（环回模式）
			RETURN_IF_FAILED(m_AudioClient->Initialize(
				AUDCLNT_SHAREMODE_SHARED,
				AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
				200000,  // 缓冲区持续时间（200毫秒）
				0,
				&m_CaptureFormat,
				nullptr));

			// 获取缓冲区大小
			RETURN_IF_FAILED(m_AudioClient->GetBufferSize(&m_BufferFrames));

			// 获取音频捕获客户端
			RETURN_IF_FAILED(m_AudioClient->GetService(IID_PPV_ARGS(&m_AudioCaptureClient)));

			// 创建异步结果对象
			RETURN_IF_FAILED(MFCreateAsyncResult(nullptr, &m_xSampleReady, nullptr, &m_SampleReadyAsyncResult));

			// 设置事件句柄
			RETURN_IF_FAILED(m_AudioClient->SetEventHandle(m_SampleReadyEvent.get()));

			// 创建WAV文件
			RETURN_IF_FAILED(CreateWAVFile());

			// 更新设备状态为已初始化
			m_DeviceState = DeviceState::Initialized;
			return S_OK;
		}());
}

// 开始捕获回调
HRESULT CLoopbackCapture::OnStartCapture(IMFAsyncResult* pResult)
{
	return SetDeviceStateErrorIfFailed([&]()->HRESULT
		{
			// 启动音频客户端
			RETURN_IF_FAILED(m_AudioClient->Start());

			// 更新设备状态为捕获中
			m_DeviceState = DeviceState::Capturing;

			// 启动写入线程
			m_bIsCapturing = true;
			m_writerThreadResult = S_OK;
			m_WriterThread = std::thread(&CLoopbackCapture::WriterThreadProc, this);

			// 将样本就绪工作项放入等待队列
			MFPutWaitingWorkItem(m_SampleReadyEvent.get(), 0, m_SampleReadyAsyncResult.get(), &m_SampleReadyKey);
			return S_OK;
		}());
}

// 异步停止捕获
HRESULT CLoopbackCapture::StopCaptureAsync()
{
	// 检查设备状态是否有效
	RETURN_HR_IF(E_NOT_VALID_STATE, (m_DeviceState != DeviceState::Capturing) && (m_DeviceState != DeviceState::Error));

	// 如果已经在停止或已停止，直接返回
	if (m_DeviceState == DeviceState::Stopping || m_DeviceState == DeviceState::Stopped)
	{
		return S_OK;
	}

	// 更新设备状态为停止中
	m_DeviceState = DeviceState::Stopping;

	// 将停止捕获工作项放入工作队列
	RETURN_IF_FAILED(MFPutWorkItem2(MFASYNC_CALLBACK_QUEUE_MULTITHREADED, 0, &m_xStopCapture, nullptr));

	// 等待捕获完全停止
	m_hCaptureStopped.wait();

	// 等待写入线程结束
	if (m_WriterThread.joinable())
	{
		m_WriterThread.join();
	}

	// 更新设备状态为已停止
	m_DeviceState = DeviceState::Stopped;

	// 返回写入线程的结果
	return m_writerThreadResult;
}

// 停止捕获回调
HRESULT CLoopbackCapture::OnStopCapture(IMFAsyncResult* pResult)
{
	// 取消样本就绪工作项
	if (0 != m_SampleReadyKey)
	{
		MFCancelWorkItem(m_SampleReadyKey);
		m_SampleReadyKey = 0;
	}

	// 停止音频客户端
	m_AudioClient->Stop();

	// 重置异步结果对象
	m_SampleReadyAsyncResult.reset();

	// 更新捕获状态
	m_bIsCapturing = false;

	// 通知写入线程
	m_QueueCV.notify_one();

	// 设置捕获停止事件
	m_hCaptureStopped.SetEvent();

	return S_OK;
}

// 样本就绪回调
HRESULT CLoopbackCapture::OnSampleReady(IMFAsyncResult* pResult)
{
	// 处理音频样本请求
	if (SUCCEEDED(OnAudioSampleRequested()))
	{
		// 如果仍在捕获中，继续等待下一个样本
		if (m_DeviceState == DeviceState::Capturing)
		{
			return MFPutWaitingWorkItem(m_SampleReadyEvent.get(), 0, m_SampleReadyAsyncResult.get(), &m_SampleReadyKey);
		}
	}
	else
	{
		// 如果处理失败，设置设备状态为错误
		m_DeviceState = DeviceState::Error;
	}
	return S_OK;
}

// 处理音频样本请求
HRESULT CLoopbackCapture::OnAudioSampleRequested()
{
	UINT32 FramesAvailable = 0;
	BYTE* Data = nullptr;
	DWORD dwCaptureFlags;
	UINT64 u64DevicePosition = 0;
	UINT64 u64QPCPosition = 0;

	// 获取临界区锁
	auto lock = m_CritSec.lock();

	// 检查设备状态
	if (m_DeviceState == DeviceState::Stopping || m_DeviceState == DeviceState::Stopped)
	{
		return S_OK;
	}

	// 处理所有可用的音频数据包
	while (SUCCEEDED(m_AudioCaptureClient->GetNextPacketSize(&FramesAvailable)) && FramesAvailable > 0)
	{
		// 计算需要捕获的字节数
		UINT32 cbBytesToCapture = FramesAvailable * m_CaptureFormat.nBlockAlign;

		// 获取音频缓冲区
		RETURN_IF_FAILED(m_AudioCaptureClient->GetBuffer(&Data, &FramesAvailable, &dwCaptureFlags, &u64DevicePosition, &u64QPCPosition));

		try
		{
			// 复制音频数据到向量
			std::vector<BYTE> audioChunk(Data, Data + cbBytesToCapture);

			// 将音频数据放入队列
			{
				std::lock_guard<std::mutex> queueLock(m_QueueMutex);
				m_AudioQueue.push(std::move(audioChunk));
			}

			// 通知写入线程有新数据
			m_QueueCV.notify_one();
		}
		catch (const std::bad_alloc&)
		{
			// 内存分配失败处理
			m_writerThreadResult = E_OUTOFMEMORY;
			StopCaptureAsync();
			break;
		}

		// 释放音频缓冲区
		m_AudioCaptureClient->ReleaseBuffer(FramesAvailable);
	}
	return S_OK;
}

// 写入线程处理函数
void CLoopbackCapture::WriterThreadProc()
{
	// 循环处理音频数据，直到捕获停止且队列为空
	while (m_bIsCapturing || !m_AudioQueue.empty())
	{
		std::vector<BYTE> audioData;

		// 从队列中获取音频数据
		{
			std::unique_lock<std::mutex> lock(m_QueueMutex);

			// 等待队列中有数据或捕获停止
			m_QueueCV.wait(lock, [this] {
				return !m_AudioQueue.empty() || !m_bIsCapturing;
				});

			// 如果队列为空，继续等待
			if (m_AudioQueue.empty())
			{
				continue;
			}

			// 获取队列前面的数据
			audioData = std::move(m_AudioQueue.front());
			m_AudioQueue.pop();
		}

		// 将音频数据写入文件
		if (!audioData.empty())
		{
			DWORD dwBytesWritten = 0;
			if (!WriteFile(m_hFile.get(), audioData.data(), static_cast<DWORD>(audioData.size()), &dwBytesWritten, NULL))
			{
				// 写入失败处理
				m_writerThreadResult = HRESULT_FROM_WIN32(GetLastError());
				m_bIsCapturing = false;
				continue;
			}

			// 更新已写入数据大小
			m_cbDataSize += dwBytesWritten;
		}
	}

	// 捕获完成后修复WAV文件头
	if (SUCCEEDED(m_writerThreadResult))
	{
		HRESULT hr = FixWAVHeader();
		if (FAILED(hr))
		{
			m_writerThreadResult = hr;
		}
	}
}