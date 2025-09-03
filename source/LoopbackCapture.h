/*
 * CLoopbackCapture��ʵ��ϵͳ��Ƶ���ز����ܣ��ɽ�ϵͳ��Ƶ��ָ�����̵���Ƶ�������WAV�ļ���
 * ��Ҫ���ܣ�
 *   - ȫ����Ƶ���񣺲�������ϵͳ��Ƶ���
 *   - �����ض����񣺲���ָ������(�����ӽ���)����Ƶ���
 *   - �첽������ͨ��Media Foundation�첽�ӿ�ʵ�ַ�����ʽ��Ƶ����
 *   - ���̴߳���ʹ�ö����̴߳�����Ƶ����д�룬�������������߳�
 *
 * ʵ��˼·��
 *   1. ��ʼ���׶Σ�������Ƶ�ӿڣ���ȡIAudioClient��IAudioCaptureClient
 *   2. ׼���׶Σ�������Ƶ��ʽ������WAV�ļ�ͷ
 *   3. ����׶Σ�������Ƶ����ͨ���첽�ص������������Ƶ����
 *   4. д��׶Σ����������Ƶ���ݷ�����У���ר���߳�д���ļ�
 *   5. ֹͣ�׶Σ�ֹͣ�����޸�WAV�ļ�ͷ���ͷ���Դ
 *
 * ʹ�õ��ļ���/�⣺
 *   - Windows Core Audio API (AudioClient.h, mmdeviceapi.h)
 *   - Media Foundation API (mfapi.h)
 *   - Windows Implementation Library (WIL) - ��COM����Դ����
 *   - C++��׼�̺߳�ͬ��ԭ��
 *
 * �ر�ע�⣺
 *   - ʹ��COM�������Ҫ��ȷ�������ü���
 *   - ���̻߳�������Ҫ��������ͬ��
 *   - WAV�ļ�ͷ��Ҫ�ڲ�����ɺ���������д����ȷ�����ݴ�С
 */

#pragma once  // ��ֹͷ�ļ��ظ�����

 // Windows��Ƶ���ͷ�ļ�
#include <AudioClient.h>    // ��Ƶ�ͻ��˽ӿ�
#include <mmdeviceapi.h>    // ��ý���豸API
#include <initguid.h>       // ��ʼ��GUID����
#include <guiddef.h>        // GUID����
#include <mfapi.h>          // Media Foundation API

// WIL�� - Windows Implementation Library����COM����Դ����
#include <wrl\implements.h> // COMʵ�ָ���
#include <wil\com.h>        // COM����ָ��
#include <wil\result.h>     // ������

// ��Ŀͨ��ͷ�ļ�
#include "Common.h"

// C++��׼��
#include <thread>            // �߳�֧��
#include <vector>            // ��̬����
#include <queue>             // ��������
#include <mutex>             // ������
#include <condition_variable> // ��������
#include <atomic>            // ԭ�Ӳ���

using namespace Microsoft::WRL;  // ʹ��WRL�����ռ�

// CLoopbackCapture������
// �̳���RuntimeClass���ṩCOM֧�֣���FtmBase��֧�������̷߳��ͣ���IActivateAudioInterfaceCompletionHandler����Ƶ�ӿڼ�����ɻص���
class CLoopbackCapture :
 public RuntimeClass< RuntimeClassFlags< ClassicCom >, FtmBase, IActivateAudioInterfaceCompletionHandler >
{
public:
 CLoopbackCapture();   // ���캯��
 ~CLoopbackCapture();  // ��������

 // ��ȡֹͣ�¼�����������ⲿ�ȴ�����ֹͣ
 HANDLE GetStopEventHandle() { return m_hCaptureStopped.get(); }

 // ��ʼȫ����Ƶ���񣨲�������ϵͳ��Ƶ��
 HRESULT StartGlobalCaptureAsync(PCWSTR outputFileName);

 // ��ʼ�����ض���Ƶ����
 HRESULT StartCaptureAsync(DWORD processId, bool includeProcessTree, PCWSTR outputFileName);

 // ֹͣ��Ƶ����
 HRESULT StopCaptureAsync();

 // ʹ�ú������첽�ص�ʵ��
 METHODASYNCCALLBACK(CLoopbackCapture, StartCapture, OnStartCapture);
 METHODASYNCCALLBACK(CLoopbackCapture, StopCapture, OnStopCapture);
 METHODASYNCCALLBACK(CLoopbackCapture, SampleReady, OnSampleReady);

 // IActivateAudioInterfaceCompletionHandler�ӿڷ���
 // ��Ƶ�ӿڼ������ʱ����
 STDMETHOD(ActivateCompleted)(IActivateAudioInterfaceAsyncOperation* operation);

private:
 // �豸״̬ö��
 enum class DeviceState
 {
  Uninitialized,  // δ��ʼ��
  Error,          // ����״̬
  Initialized,    // �ѳ�ʼ��
  Starting,       // ��������
  Capturing,      // ���ڲ���
  Stopping,       // ����ֹͣ
  Stopped,        // ��ֹͣ
 };

 // �첽�ص�������
 HRESULT OnStartCapture(IMFAsyncResult* pResult);  // ��ʼ����ص�
 HRESULT OnStopCapture(IMFAsyncResult* pResult);   // ֹͣ����ص�
 HRESULT OnSampleReady(IMFAsyncResult* pResult);   // ���������ص�

 // ��ʼ�����ز���
 HRESULT InitializeLoopbackCapture();

 // ����WAV�ļ���д���ļ�ͷ
 HRESULT CreateWAVFile();

 // �޸�WAV�ļ�ͷ���ڲ�����ɺ�д����ȷ�����ݴ�С��
 HRESULT FixWAVHeader();

 // ������Ƶ��������
 HRESULT OnAudioSampleRequested();

 // ����ָ�����̵���Ƶ�ӿ�
 HRESULT ActivateAudioInterface(DWORD processId, bool includeProcessTree);

 // �������ʧ���������豸״̬Ϊ����
 HRESULT SetDeviceStateErrorIfFailed(HRESULT hr);

 // ����ȫ����Ƶ�ӿ�
 HRESULT ActivateAudioInterfaceGlobal();

 // д���̴߳�����
 void WriterThreadProc();

 // ��Ա����

 wil::com_ptr_nothrow<IAudioClient> m_AudioClient;        // ��Ƶ�ͻ��˽ӿ�
 WAVEFORMATEX m_CaptureFormat{};                         // �������Ƶ��ʽ
 UINT32 m_BufferFrames = 0;                              // ������֡��
 wil::com_ptr_nothrow<IAudioCaptureClient> m_AudioCaptureClient;  // ��Ƶ����ͻ��˽ӿ�
 wil::com_ptr_nothrow<IMFAsyncResult> m_SampleReadyAsyncResult;   // ���������첽���

 wil::unique_event_nothrow m_SampleReadyEvent;           // ���������¼�
 MFWORKITEM_KEY m_SampleReadyKey = 0;                    // Media Foundation�������
 wil::unique_hfile m_hFile;                              // ����ļ����
 wil::critical_section m_CritSec;                        // �ٽ���������ͬ��
 DWORD m_dwQueueID = 0;                                  // �첽����ID
 DWORD m_cbHeaderSize = 0;                               // WAV�ļ�ͷ��С
 DWORD m_cbDataSize = 0;                                 // ��Ƶ���ݴ�С

 PCWSTR m_outputFileName = nullptr;                      // ����ļ���
 HRESULT m_activateResult = E_UNEXPECTED;                // ����������

 DeviceState m_DeviceState{ DeviceState::Uninitialized }; // ��ǰ�豸״̬
 wil::unique_event_nothrow m_hActivateCompleted;         // ��������¼�
 wil::unique_event_nothrow m_hCaptureStopped;            // ����ֹͣ�¼�

 std::thread m_WriterThread;                             // ��Ƶ����д���߳�
 std::queue<std::vector<BYTE>> m_AudioQueue;             // ��Ƶ���ݶ���
 std::mutex m_QueueMutex;                                // ���л�����
 std::condition_variable m_QueueCV;                      // ������������
 std::atomic<bool> m_bIsCapturing;                       // ����״̬��־��ԭ�Ӳ�����
 std::atomic<HRESULT> m_writerThreadResult;              // д���߳̽����ԭ�Ӳ�����
};