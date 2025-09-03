/*
 * 该宏定义用于简化IMFAsyncCallback接口的实现，允许父类将异步回调委托给指定的成员函数。
 * 主要功能：创建一个内嵌的IMFAsyncCallback实现类，该类将Invoke调用转发给父类的指定方法。
 * 实现思路：
 *   1. 通过宏生成一个嵌套类，该类继承IMFAsyncCallback并实现其接口方法
 *   2. 使用offsetof计算父类指针，建立回调类与父类的关联
 *   3. 通过函数指针将Invoke调用转发给父类的指定方法
 * 注意：
 *   - 使用了Windows Media Foundation库（mfidl.h等）
 *   - 依赖offsetof宏，要求父类必须是标准布局类型
 *   - 会在父类中声明一个名为m_x##AsyncCallback的成员变量
 */

#pragma once  // 防止头文件重复包含

#include <mfidl.h>   // Windows Media Foundation接口定义
#include <mfapi.h>   // Media Foundation API
#include <mfobjects.h> // Media Foundation对象

#ifndef METHODASYNCCALLBACK  // 防止宏重复定义
#define METHODASYNCCALLBACK(Parent, AsyncCallback, pfnCallback) \
/* 定义回调类，类名为Callback##AsyncCallback */ \
class Callback##AsyncCallback :\
    public IMFAsyncCallback \
{ \
public: \
    /* 构造函数：通过offsetof计算父对象地址并初始化队列ID */ \
    Callback##AsyncCallback() : \
        _parent(((Parent*)((BYTE*)this - offsetof(Parent, m_x##AsyncCallback)))), \
        _dwQueueID( MFASYNC_CALLBACK_QUEUE_MULTITHREADED ) \
    { \
    } \
\
    /* 委托引用计数到父对象 */ \
    STDMETHOD_( ULONG, AddRef )() \
    { \
        return _parent->AddRef(); \
    } \
    STDMETHOD_( ULONG, Release )() \
    { \
        return _parent->Release(); \
    } \
    /* 查询接口：支持IMFAsyncCallback和IUnknown */ \
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
    /* 获取回调参数：设置标志和队列ID */ \
    STDMETHOD( GetParameters )( \
        /* [out] */ __RPC__out DWORD *pdwFlags, \
        /* [out] */ __RPC__out DWORD *pdwQueue) \
    { \
        *pdwFlags = 0; \
        *pdwQueue = _dwQueueID; \
        return S_OK; \
    } \
    /* 回调执行：转发到父类的指定方法 */ \
    STDMETHOD( Invoke )( /* [out] */ __RPC__out IMFAsyncResult * pResult ) \
    { \
        _parent->pfnCallback( pResult ); \
        return S_OK; \
    } \
    /* 设置队列ID */ \
    void SetQueueID( DWORD dwQueueID ) { _dwQueueID = dwQueueID; } \
\
protected: \
    Parent* _parent;       /* 指向父对象的指针 */ \
    DWORD   _dwQueueID;    /* 回调队列标识 */ \
           \
} m_x##AsyncCallback;  // 在父类中声明的成员变量
#endif