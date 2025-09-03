/*
 * CLoopbackCapture类实现系统音频环回捕获功能，可将系统音频或指定进程的音频输出捕获到WAV文件。
 * 主要功能：
 *   - 全局音频捕获：捕获所有系统音频输出
 *   - 进程特定捕获：捕获指定进程(及其子进程)的音频输出
 *   - 异步操作：通过Media Foundation异步接口实现非阻塞式音频捕获
 *   - 多线程处理：使用独立线程处理音频数据写入，避免阻塞捕获线程
 *
 * 实现思路：
 *   1. 初始化阶段：激活音频接口，获取IAudioClient和IAudioCaptureClient
 *   2. 准备阶段：设置音频格式，创建WAV文件头
 *   3. 捕获阶段：启动音频捕获，通过异步回调处理就绪的音频样本
 *   4. 写入阶段：将捕获的音频数据放入队列，由专门线程写入文件
 *   5. 停止阶段：停止捕获，修复WAV文件头，释放资源
 *
 * 使用到的技术/库：
 *   - Windows Core Audio API (AudioClient.h, mmdeviceapi.h)
 *   - Media Foundation API (mfapi.h)
 *   - Windows Implementation Library (WIL) - 简化COM和资源管理
 *   - C++标准线程和同步原语
 *
 * 特别注意：
 *   - 使用COM组件，需要正确管理引用计数
 *   - 多线程环境下需要谨慎处理同步
 *   - WAV文件头需要在捕获完成后修正，以写入正确的数据大小
 */

#pragma once  // 防止头文件重复包含

 // Windows音频相关头文件
#include <AudioClient.h>    // 音频客户端接口
#include <mmdeviceapi.h>    // 多媒体设备API
#include <initguid.h>       // 初始化GUID定义
#include <guiddef.h>        // GUID定义
#include <mfapi.h>          // Media Foundation API

// WIL库 - Windows Implementation Library，简化COM和资源管理
#include <wrl\implements.h> // COM实现辅助
#include <wil\com.h>        // COM智能指针
#include <wil\result.h>     // 错误处理

// 项目通用头文件
#include "Common.h"

// C++标准库
#include <thread>            // 线程支持
#include <vector>            // 动态数组
#include <queue>             // 队列容器
#include <mutex>             // 互斥锁
#include <condition_variable> // 条件变量
#include <atomic>            // 原子操作

using namespace Microsoft::WRL;  // 使用WRL命名空间

// CLoopbackCapture类声明
// 继承自RuntimeClass（提供COM支持）、FtmBase（支持自由线程封送）和IActivateAudioInterfaceCompletionHandler（音频接口激活完成回调）
class CLoopbackCapture :
 public RuntimeClass< RuntimeClassFlags< ClassicCom >, FtmBase, IActivateAudioInterfaceCompletionHandler >
{
public:
 CLoopbackCapture();   // 构造函数
 ~CLoopbackCapture();  // 析构函数

 // 获取停止事件句柄，用于外部等待捕获停止
 HANDLE GetStopEventHandle() { return m_hCaptureStopped.get(); }

 // 开始全局音频捕获（捕获所有系统音频）
 HRESULT StartGlobalCaptureAsync(PCWSTR outputFileName);

 // 开始进程特定音频捕获
 HRESULT StartCaptureAsync(DWORD processId, bool includeProcessTree, PCWSTR outputFileName);

 // 停止音频捕获
 HRESULT StopCaptureAsync();

 // 使用宏生成异步回调实现
 METHODASYNCCALLBACK(CLoopbackCapture, StartCapture, OnStartCapture);
 METHODASYNCCALLBACK(CLoopbackCapture, StopCapture, OnStopCapture);
 METHODASYNCCALLBACK(CLoopbackCapture, SampleReady, OnSampleReady);

 // IActivateAudioInterfaceCompletionHandler接口方法
 // 音频接口激活完成时调用
 STDMETHOD(ActivateCompleted)(IActivateAudioInterfaceAsyncOperation* operation);

private:
 // 设备状态枚举
 enum class DeviceState
 {
  Uninitialized,  // 未初始化
  Error,          // 错误状态
  Initialized,    // 已初始化
  Starting,       // 正在启动
  Capturing,      // 正在捕获
  Stopping,       // 正在停止
  Stopped,        // 已停止
 };

 // 异步回调处理函数
 HRESULT OnStartCapture(IMFAsyncResult* pResult);  // 开始捕获回调
 HRESULT OnStopCapture(IMFAsyncResult* pResult);   // 停止捕获回调
 HRESULT OnSampleReady(IMFAsyncResult* pResult);   // 样本就绪回调

 // 初始化环回捕获
 HRESULT InitializeLoopbackCapture();

 // 创建WAV文件并写入文件头
 HRESULT CreateWAVFile();

 // 修复WAV文件头（在捕获完成后写入正确的数据大小）
 HRESULT FixWAVHeader();

 // 处理音频样本请求
 HRESULT OnAudioSampleRequested();

 // 激活指定进程的音频接口
 HRESULT ActivateAudioInterface(DWORD processId, bool includeProcessTree);

 // 如果操作失败则设置设备状态为错误
 HRESULT SetDeviceStateErrorIfFailed(HRESULT hr);

 // 激活全局音频接口
 HRESULT ActivateAudioInterfaceGlobal();

 // 写入线程处理函数
 void WriterThreadProc();

 // 成员变量

 wil::com_ptr_nothrow<IAudioClient> m_AudioClient;        // 音频客户端接口
 WAVEFORMATEX m_CaptureFormat{};                         // 捕获的音频格式
 UINT32 m_BufferFrames = 0;                              // 缓冲区帧数
 wil::com_ptr_nothrow<IAudioCaptureClient> m_AudioCaptureClient;  // 音频捕获客户端接口
 wil::com_ptr_nothrow<IMFAsyncResult> m_SampleReadyAsyncResult;   // 样本就绪异步结果

 wil::unique_event_nothrow m_SampleReadyEvent;           // 样本就绪事件
 MFWORKITEM_KEY m_SampleReadyKey = 0;                    // Media Foundation工作项键
 wil::unique_hfile m_hFile;                              // 输出文件句柄
 wil::critical_section m_CritSec;                        // 临界区，用于同步
 DWORD m_dwQueueID = 0;                                  // 异步队列ID
 DWORD m_cbHeaderSize = 0;                               // WAV文件头大小
 DWORD m_cbDataSize = 0;                                 // 音频数据大小

 PCWSTR m_outputFileName = nullptr;                      // 输出文件名
 HRESULT m_activateResult = E_UNEXPECTED;                // 激活操作结果

 DeviceState m_DeviceState{ DeviceState::Uninitialized }; // 当前设备状态
 wil::unique_event_nothrow m_hActivateCompleted;         // 激活完成事件
 wil::unique_event_nothrow m_hCaptureStopped;            // 捕获停止事件

 std::thread m_WriterThread;                             // 音频数据写入线程
 std::queue<std::vector<BYTE>> m_AudioQueue;             // 音频数据队列
 std::mutex m_QueueMutex;                                // 队列互斥锁
 std::condition_variable m_QueueCV;                      // 队列条件变量
 std::atomic<bool> m_bIsCapturing;                       // 捕获状态标志（原子操作）
 std::atomic<HRESULT> m_writerThreadResult;              // 写入线程结果（原子操作）
};