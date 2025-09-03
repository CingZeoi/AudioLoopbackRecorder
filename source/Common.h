/*
 * �ú궨�����ڼ�IMFAsyncCallback�ӿڵ�ʵ�֣������ཫ�첽�ص�ί�и�ָ���ĳ�Ա������
 * ��Ҫ���ܣ�����һ����Ƕ��IMFAsyncCallbackʵ���࣬���ཫInvoke����ת���������ָ��������
 * ʵ��˼·��
 *   1. ͨ��������һ��Ƕ���࣬����̳�IMFAsyncCallback��ʵ����ӿڷ���
 *   2. ʹ��offsetof���㸸��ָ�룬�����ص����븸��Ĺ���
 *   3. ͨ������ָ�뽫Invoke����ת���������ָ������
 * ע�⣺
 *   - ʹ����Windows Media Foundation�⣨mfidl.h�ȣ�
 *   - ����offsetof�꣬Ҫ��������Ǳ�׼��������
 *   - ���ڸ���������һ����Ϊm_x##AsyncCallback�ĳ�Ա����
 */

#pragma once  // ��ֹͷ�ļ��ظ�����

#include <mfidl.h>   // Windows Media Foundation�ӿڶ���
#include <mfapi.h>   // Media Foundation API
#include <mfobjects.h> // Media Foundation����

#ifndef METHODASYNCCALLBACK  // ��ֹ���ظ�����
#define METHODASYNCCALLBACK(Parent, AsyncCallback, pfnCallback) \
/* ����ص��࣬����ΪCallback##AsyncCallback */ \
class Callback##AsyncCallback :\
    public IMFAsyncCallback \
{ \
public: \
    /* ���캯����ͨ��offsetof���㸸�����ַ����ʼ������ID */ \
    Callback##AsyncCallback() : \
        _parent(((Parent*)((BYTE*)this - offsetof(Parent, m_x##AsyncCallback)))), \
        _dwQueueID( MFASYNC_CALLBACK_QUEUE_MULTITHREADED ) \
    { \
    } \
\
    /* ί�����ü����������� */ \
    STDMETHOD_( ULONG, AddRef )() \
    { \
        return _parent->AddRef(); \
    } \
    STDMETHOD_( ULONG, Release )() \
    { \
        return _parent->Release(); \
    } \
    /* ��ѯ�ӿڣ�֧��IMFAsyncCallback��IUnknown */ \
    STDMETHOD( QueryInterface )( REFIID riid, void **ppvObject ) \
    { \
        if (riid == IID_IMFAsyncCallback || riid == IID_IUnknown) \
        { \
            (*ppvObject) = this; \
            AddRef(); \
            return S_OK; \
        } \
        *ppvObject = NULL; \
        return E_NOINTERFACE; \
    } \
    /* ��ȡ�ص����������ñ�־�Ͷ���ID */ \
    STDMETHOD( GetParameters )( \
        /* [out] */ __RPC__out DWORD *pdwFlags, \
        /* [out] */ __RPC__out DWORD *pdwQueue) \
    { \
        *pdwFlags = 0; \
        *pdwQueue = _dwQueueID; \
        return S_OK; \
    } \
    /* �ص�ִ�У�ת���������ָ������ */ \
    STDMETHOD( Invoke )( /* [out] */ __RPC__out IMFAsyncResult * pResult ) \
    { \
        _parent->pfnCallback( pResult ); \
        return S_OK; \
    } \
    /* ���ö���ID */ \
    void SetQueueID( DWORD dwQueueID ) { _dwQueueID = dwQueueID; } \
\
protected: \
    Parent* _parent;       /* ָ�򸸶����ָ�� */ \
    DWORD   _dwQueueID;    /* �ص����б�ʶ */ \
           \
} m_x##AsyncCallback;  // �ڸ����������ĳ�Ա����
#endif