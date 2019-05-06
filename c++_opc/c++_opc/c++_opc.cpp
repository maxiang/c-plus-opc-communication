// c++_opc.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <atlbase.h>
#include <objbase.h>
#include <iostream>
#include <windows.h>
#include<vector>
#include "Opcda_i.c"
#include "Opcda.h"
#include "Opccomn_i.c"
#include "Opccomn.h"
#include "OpcError.h"
#include "OPCClient.h"

using namespace std;

//VT_EMPTY= 0,VT_NULL= 1,VT_I2= 2,VT_I4	= 3,VT_R4	= 4,VT_R8	= 5,VT_CY	= 6,VT_DATE	= 7,VT_BSTR	= 8,VT_DISPATCH	= 9,VT_ERROR= 10,VT_BOOL= 11,VT_VARIANT	= 12,VT_UNKNOWN	= 13,
//VT_DECIMAL= 14,VT_I1	= 16,VT_UI1	= 17,VT_UI2	= 18,VT_UI4	= 19,VT_I8	= 20,VT_UI8	= 21,VT_INT	= 22,VT_UINT= 23,VT_VOID= 24,VT_HRESULT	= 25,VT_PTR	= 26,VT_SAFEARRAY= 27,
//VT_CARRAY	= 28,VT_USERDEFINED	= 29,VT_LPSTR	= 30,VT_LPWSTR	= 31,VT_RECORD	= 36,VT_INT_PTR	= 37,VT_UINT_PTR= 38,VT_FILETIME= 64,VT_BLOB= 65,VT_STREAM= 66,VT_STORAGE= 67,
//VT_STREAMED_OBJECT= 68,VT_STORED_OBJECT= 69,VT_BLOB_OBJECT= 70,VT_CF	= 71,VT_CLSID= 72,VT_VERSIONED_STREAM= 73,VT_BSTR_BLOB	= 0xfff,VT_VECTOR= 0x1000,VT_ARRAY	= 0x2000,
//VT_BYREF	= 0x4000,VT_RESERVED= 0x8000,VT_ILLEGAL	= 0xffff,VT_ILLEGALMASKED	= 0xfff,VT_TYPEMASK	= 0xfff


//#define OPC_SERVER_NAME  L"Matrikon.OPC.Simulation.1" 
#define OPC_SERVER_NAME  L"KingView.View.1"
#define OPC_GROUP_NAME_r   "Read"
#define OPC_GROUP_NAME_w   "Write"
#define VT 0//VT_R4
#define XVAL fltVal
#define OPC_REQ_UPDATE_RATE 500 /* (msec) Requested Update Rate*/

static wchar_t* m_wchar = NULL; 
const string itemsOPC_read = ("sec.Value,fou.Value");
const string itemsOPC_write = ("first.Value,thi.Value");


void main(void)
{
	IOPCServer* pIOPCServer = NULL;   
	IOPCItemMgt* pIOPCItemMgt_r = NULL; 
	IOPCItemMgt* pIOPCItemMgt_w = NULL; 
	OPCHANDLE hServerGroup_r; // OPC Group句柄
	OPCHANDLE hServerGroup_w; // OPC Group句柄

	vector<string> itemNames_r = split(itemsOPC_read,",");
	vector<string> itemNames_w = split(itemsOPC_write,",");
	OPCHANDLE *hItems_r = new OPCHANDLE[itemNames_r.size()];
	OPCHANDLE *hItems_w = new OPCHANDLE[itemNames_w.size()];

	//初始化COM库。若成功，则返回值等于S_OK
	HRESULT  r1;
	IMalloc *g_pIMalloc = NULL;
	r1 = CoInitialize(NULL);
	r1 = CoGetMalloc(MEMCTX_TASK, &g_pIMalloc);//得到一个指向COM内存管理接口的指针
	//实例化IOPCServer 并得到其指针
	pIOPCServer = InstantiateServer(OPC_SERVER_NAME);

	AddTheGroup(pIOPCServer, pIOPCItemMgt_r, hServerGroup_r, StringToWchar(OPC_GROUP_NAME_r));
	AddTheGroup(pIOPCServer, pIOPCItemMgt_w, hServerGroup_w, StringToWchar(OPC_GROUP_NAME_w));

	AddItemsToGroup(pIOPCItemMgt_r, hItems_r, itemNames_r);
	AddItemsToGroup(pIOPCItemMgt_w, hItems_w, itemNames_w);

	VARIANT varValue_r; //to stor the read value
	VariantInit(&varValue_r);
	VARIANT varValue_w[2];
	VariantInit(varValue_w);
	varValue_w[0].vt = VT_R4;//数据类型为R4
	//varValue_w[0].fltVal = 234.0;
	varValue_w[1].vt = VT_R4;//数据类型为R4
	//varValue_w[1].fltVal = 567.0;
	int i = 0;
	while(1)
	{
		ReadItems(pIOPCItemMgt_r, hItems_r, itemNames_r.size(), varValue_r);
		//WriteItem(pIOPCItemMgt_w, hItems_w,varValue_w);
		varValue_w[0].fltVal = i;//浮点型数据的具体内容
		varValue_w[1].fltVal = sqrt(i);
		i++;
		WriteItems(pIOPCItemMgt_w, hItems_w,itemNames_w.size(), varValue_w);
		Sleep(1000);
		if(i == 100)
			i = 0;
	}

	RemoveItems(pIOPCItemMgt_r, hItems_r, itemNames_r.size());
    RemoveGroup(pIOPCServer, hServerGroup_r);

	pIOPCItemMgt_r->Release();
	pIOPCServer->Release();

	//关闭当前线程的COM
	CoUninitialize();
	/*delete[] hItems;*/

}




/****************************************************
将OPCServer的接口IOPCServer 实例化，返回这个实例的指针。
参数ServerName为IOPCServer的名字
*****************************************************/
IOPCServer* InstantiateServer(wchar_t *ServerName)
{
	CLSID CLSID_OPCServer;
	HRESULT hr;
	IOPCServer* m_pIOPCServer = NULL;
	IUnknown *pUNK;

	// 从OPC Server Name得到CLSID
	//hr = CLSIDFromString(ServerName, &CLSID_OPCServer);
	 hr = CLSIDFromProgID(ServerName, &CLSID_OPCServer);
	_ASSERT(!FAILED(hr));

	//LONG cmq = 1;//指明要查询接口的个数
	//MULTI_QI queue[1] ={{&IID_IOPCServer, NULL, 0}};
	//MULTI_QI   queue[2];//用于接收查询到的接口，可以为数组，以接收多个接口。

	//创建一个Com对象
	//hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_SERVER,
	//	/*&CoServerInfo*/NULL, cmq, queue);
	//hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_LOCAL_SERVER,
	//	/*&CoServerInfo*/NULL, cmq, queue);
	//hr = CoCreateInstance (CLSID_OPCServer, NULL, CLSCTX_LOCAL_SERVER ,IID_IOPCServer, (void**)&m_pIOPCServer);
	hr = CoCreateInstance (CLSID_OPCServer, NULL, CLSCTX_LOCAL_SERVER ,IID_IUnknown, (void**)&pUNK);
	_ASSERT(!hr);

	if (pUNK)
	{
    // Request an IOPCServer interface from the object.
    //
		hr = pUNK->QueryInterface(IID_IOPCServer, (void**)&m_pIOPCServer);
		_ASSERT(!hr);
		pUNK->Release();
		pUNK = NULL;
	}
	//return(IOPCServer*) queue[0].pItf;
	return m_pIOPCServer;
}
/***********************************************
以参数groupName作为组名，将该组添加到Server中，
若想更改刷新率，增加一个函数参数即可
************************************************/
void AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt, 
				 OPCHANDLE& hServerGroup, wchar_t groupName[])
{
	DWORD dwUpdateRate;
	OPCHANDLE hClientGroup = 1;
	LONG TimBias = 0;
	FLOAT PercDeadband = 0.0;
	DWORD dwLCID = 0x409;

  //  HRESULT hr = pIOPCServer->AddGroup(/*szName*/ groupName,
		///*bActive*/ FALSE,
		///*dwRequestedUpdateRate*/ dwUpdateRate,//服务器向客户程序提交数据变化的刷新速率
		///*hClientGroup*/ hClientGroup,
		///*pTimeBias*/ 0,
		///*pPercentDeadband*/ 0,
		///*dwLCID*/0,
		///*phServerGroup*/&hServerGroup,
		//&dwUpdateRate,
		///*riid*/ IID_IOPCItemMgt,
		///*ppUnk*/ (IUnknown**) &pIOPCItemMgt);
	  HRESULT hr = pIOPCServer->AddGroup(/*szName*/ groupName,
		/*bActive*/ TRUE,
		/*dwRequestedUpdateRate*/ OPC_REQ_UPDATE_RATE,//服务器向客户程序提交数据变化的刷新速率
		/*hClientGroup*/ hClientGroup,
		/*pTimeBias*/ &TimBias,
		/*pPercentDeadband*/ &PercDeadband,
		/*dwLCID*/dwLCID,
		/*phServerGroup*/&hServerGroup,
		&dwUpdateRate,
		/*riid*/ IID_IOPCItemMgt,
		/*ppUnk*/ (LPUNKNOWN*) &pIOPCItemMgt);
	_ASSERT(!FAILED(hr));
}

/***************************************************
将一个Item加到组中，Item的ID为itemId
OPC支持批量添加Item，如果想批量添加则在通过pIOPCItemMgt
调用AddItems前将ItemArray数组填充成你想要的个数，
将数组元素的个数以及数组名传给AddItems即可
****************************************************/
void AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, wchar_t itemId[])
{
	HRESULT hr;
	//OPCITEMDEF  *ItemArray = NULL;
	OPCITEMDEF  *ItemArray;
	WCHAR g_wstrTagPrefix[100] = L"";
	ItemArray = new OPCITEMDEF[1];

	//OPCITEMDEF ItemArray[1] =
	//{{
	///*szAccessPath*/ L"",
	///*szItemID*/ itemId,
	///*bActive*/ TRUE,
	///*hClient*/ 1,
	///*dwBlobSize*/ 0,
	///*pBlob*/ NULL,
	///*vtRequestedDataType*/ VT,
	///*wReserved*/0
	//}};
	ItemArray[0].szAccessPath = g_wstrTagPrefix;   

	ItemArray[0].szItemID = itemId;  // 影响数据类型
	ItemArray[0].bActive = TRUE;   
	ItemArray[0].hClient = 1;
	ItemArray[0].dwBlobSize = 0;
	ItemArray[0].pBlob = NULL;
	ItemArray[0].vtRequestedDataType = 0; 

	//Add Result:
	OPCITEMRESULT* pAddResult=NULL;
	HRESULT* pErrors = NULL;

	hr = pIOPCItemMgt->AddItems(1, ItemArray, &pAddResult, &pErrors);
	_ASSERT(!hr);

	hServerItem = pAddResult[0].hServer;

	//释放Server分配的内存
	CoTaskMemFree(pAddResult->pBlob);
	CoTaskMemFree(pAddResult);
	pAddResult = NULL;

	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

/*************************************************
最简单的批量添加元素到组中的实现，这种效率不够好，但是
代码简单，逻辑简单，如果想高效率的添加多个Item，则可
按照117行的注释来实现。
1.参数hItems为Item句柄数组，传出型参数
2.itemNames为Item名字的链表
*************************************************/
void AddItemsToGroup( IOPCItemMgt* pIOPCItemMgt, OPCHANDLE* hItems, vector<string>& itemNames )
{
	string str;
	//string jiange = ".";
	for(int i = 0; i < itemNames.size(); i++)
	{
		//将itemId拼接起来
		//str = string("%1%2%3%4%5%6%7").arg(PLC_HEAD, ".", PLC_NAME, ".", OPC_GROUP_NAME, ".", itemNames[i]);
		/*str += string(OPC_GROUP_NAME);
		str += jiange;
		str += itemNames[i];*/
		str = itemNames[i];
		AddTheItem(pIOPCItemMgt, hItems[i], StringToWchar(str));
	}
}

/*************************************************
1.参数hServerItem为item句柄，以此来读取item
2.Item的Value 保存在varValue中传出
注：通过pIOPCSyncIO指针调用Read方法时可以批量读取Item
**************************************************/
void ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue)
{
	OPCITEMSTATE* pValue = NULL;

	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);

	HRESULT* pErrors = NULL; 
	HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);
	_ASSERT(!hr);
	_ASSERT(pValue!=NULL);

	varValue = pValue[0].vDataValue;//得到的数据

	CoTaskMemFree(pErrors);
	pErrors = NULL;

	CoTaskMemFree(pValue);
	pValue = NULL;

	pIOPCSyncIO->Release();
}

/*******************************************************
批量读取Item的方法。哈哈，这里我偷懒了，158行注释也说了这一点
1.hItems 为Item句柄指针（传入的是句柄数组名）
2.itemCount句柄数组元素的个数
********************************************************/
void ReadItems(IOPCItemMgt* pItemMgt, OPCHANDLE* hItems, int itemCount, VARIANT& varValue)
{
	for(int i = 0; i < itemCount; i++)
	{
		ReadItem(pItemMgt, hItems[i], varValue);
		cout << "Read value: " << varValue.XVAL << endl;
	}
}

/**********************************************************
写Item
***********************************************************/
void WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue)
{
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**) &pIOPCSyncIO);

	HRESULT* pErrors; 
	HRESULT hr = pIOPCSyncIO->Write(1,&hServerItem,&varValue,&pErrors);
	_ASSERT(!hr);

	CoTaskMemFree(pErrors);
	pErrors = NULL;

	pIOPCSyncIO->Release();
}
//批量写Item的方法
void WriteItems(IOPCItemMgt* pItemMgt, OPCHANDLE* hItems, int itemCount, VARIANT* varValue)
{
	for(int i = 0; i < itemCount; i++)
	{
		WriteItem(pItemMgt, hItems[i], varValue[i]);
	}
}

/**********************************************************
删除一个Item。IOPCItemMgt::RemoveItems 具有删除多个Item的功能
***********************************************************/
void RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE hServerItem)
{
	OPCHANDLE hServerArray[1];
	hServerArray[0] = hServerItem;
	
	HRESULT* pErrors; 
	HRESULT hr = pIOPCItemMgt->RemoveItems(1, hServerArray, &pErrors);
	_ASSERT(!hr);

	CoTaskMemFree(pErrors);
	pErrors = NULL;
}


/*******************************************************************
删除多个Item的方法。这里还是偷懒了，可以直接通过IOPCItemMgt::RemoveItems
批量删除Item
********************************************************************/
void RemoveItems( IOPCItemMgt* pIOPCItemMgt, OPCHANDLE* hItems, int itemCount)
{
	for(int i = 0; i < itemCount; i++)
	{
		RemoveItem(pIOPCItemMgt, hItems[i]);
	}
}

/***********************************************
移除OPC组
************************************************/
void RemoveGroup (IOPCServer* pIOPCServer, OPCHANDLE hServerGroup)
{
	HRESULT hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);
	_ASSERT(!hr);
}



/***********************************************
	释放动态分配的内存
************************************************/
static void Release()
{
	if(m_wchar)
	{
		delete m_wchar;
		m_wchar = NULL;
	}
}

/********************************************
        char* 到wchar_t*的转换函数
*********************************************/
static wchar_t* CharToWchar(const char* c)  
{  
	Release();  
//	printf("%s\n", c);
	int len = MultiByteToWideChar(CP_ACP,0,c,strlen(c),NULL,0);  
	m_wchar=new wchar_t[len+1];  
	MultiByteToWideChar(CP_ACP,0,c,strlen(c),m_wchar,len);  
	m_wchar[len]='\0';  
	return m_wchar;  
} 

/**********************************************
        std string转换成wchar_t*的方法
***********************************************/
static wchar_t* StringToWchar(const string& s)  
{  
	const char* p=s.c_str();  
	return CharToWchar(p);  
} 

//字符串分隔函数
vector<string> split(string str, string pattern)
{
    string::size_type pos;
    vector<string> result;

    str += pattern;//扩展字符串以方便操作
    int size = str.size();

    for (int i = 0; i<size; i++) {
        pos = str.find(pattern, i);
        if ( pos < size) {
            std::string s = str.substr(i, pos - i);
            result.push_back(s);
            i = pos + pattern.size() - 1;
        }
    }
    return result;
}

