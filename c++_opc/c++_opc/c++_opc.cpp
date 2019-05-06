// c++_opc.cpp : �������̨Ӧ�ó������ڵ㡣
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
	OPCHANDLE hServerGroup_r; // OPC Group���
	OPCHANDLE hServerGroup_w; // OPC Group���

	vector<string> itemNames_r = split(itemsOPC_read,",");
	vector<string> itemNames_w = split(itemsOPC_write,",");
	OPCHANDLE *hItems_r = new OPCHANDLE[itemNames_r.size()];
	OPCHANDLE *hItems_w = new OPCHANDLE[itemNames_w.size()];

	//��ʼ��COM�⡣���ɹ����򷵻�ֵ����S_OK
	HRESULT  r1;
	IMalloc *g_pIMalloc = NULL;
	r1 = CoInitialize(NULL);
	r1 = CoGetMalloc(MEMCTX_TASK, &g_pIMalloc);//�õ�һ��ָ��COM�ڴ����ӿڵ�ָ��
	//ʵ����IOPCServer ���õ���ָ��
	pIOPCServer = InstantiateServer(OPC_SERVER_NAME);

	AddTheGroup(pIOPCServer, pIOPCItemMgt_r, hServerGroup_r, StringToWchar(OPC_GROUP_NAME_r));
	AddTheGroup(pIOPCServer, pIOPCItemMgt_w, hServerGroup_w, StringToWchar(OPC_GROUP_NAME_w));

	AddItemsToGroup(pIOPCItemMgt_r, hItems_r, itemNames_r);
	AddItemsToGroup(pIOPCItemMgt_w, hItems_w, itemNames_w);

	VARIANT varValue_r; //to stor the read value
	VariantInit(&varValue_r);
	VARIANT varValue_w[2];
	VariantInit(varValue_w);
	varValue_w[0].vt = VT_R4;//��������ΪR4
	//varValue_w[0].fltVal = 234.0;
	varValue_w[1].vt = VT_R4;//��������ΪR4
	//varValue_w[1].fltVal = 567.0;
	int i = 0;
	while(1)
	{
		ReadItems(pIOPCItemMgt_r, hItems_r, itemNames_r.size(), varValue_r);
		//WriteItem(pIOPCItemMgt_w, hItems_w,varValue_w);
		varValue_w[0].fltVal = i;//���������ݵľ�������
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

	//�رյ�ǰ�̵߳�COM
	CoUninitialize();
	/*delete[] hItems;*/

}




/****************************************************
��OPCServer�Ľӿ�IOPCServer ʵ�������������ʵ����ָ�롣
����ServerNameΪIOPCServer������
*****************************************************/
IOPCServer* InstantiateServer(wchar_t *ServerName)
{
	CLSID CLSID_OPCServer;
	HRESULT hr;
	IOPCServer* m_pIOPCServer = NULL;
	IUnknown *pUNK;

	// ��OPC Server Name�õ�CLSID
	//hr = CLSIDFromString(ServerName, &CLSID_OPCServer);
	 hr = CLSIDFromProgID(ServerName, &CLSID_OPCServer);
	_ASSERT(!FAILED(hr));

	//LONG cmq = 1;//ָ��Ҫ��ѯ�ӿڵĸ���
	//MULTI_QI queue[1] ={{&IID_IOPCServer, NULL, 0}};
	//MULTI_QI   queue[2];//���ڽ��ղ�ѯ���Ľӿڣ�����Ϊ���飬�Խ��ն���ӿڡ�

	//����һ��Com����
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
�Բ���groupName��Ϊ��������������ӵ�Server�У�
�������ˢ���ʣ�����һ��������������
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
		///*dwRequestedUpdateRate*/ dwUpdateRate,//��������ͻ������ύ���ݱ仯��ˢ������
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
		/*dwRequestedUpdateRate*/ OPC_REQ_UPDATE_RATE,//��������ͻ������ύ���ݱ仯��ˢ������
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
��һ��Item�ӵ����У�Item��IDΪitemId
OPC֧���������Item������������������ͨ��pIOPCItemMgt
����AddItemsǰ��ItemArray������������Ҫ�ĸ�����
������Ԫ�صĸ����Լ�����������AddItems����
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

	ItemArray[0].szItemID = itemId;  // Ӱ����������
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

	//�ͷ�Server������ڴ�
	CoTaskMemFree(pAddResult->pBlob);
	CoTaskMemFree(pAddResult);
	pAddResult = NULL;

	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

/*************************************************
��򵥵��������Ԫ�ص����е�ʵ�֣�����Ч�ʲ����ã�����
����򵥣��߼��򵥣�������Ч�ʵ���Ӷ��Item�����
����117�е�ע����ʵ�֡�
1.����hItemsΪItem������飬�����Ͳ���
2.itemNamesΪItem���ֵ�����
*************************************************/
void AddItemsToGroup( IOPCItemMgt* pIOPCItemMgt, OPCHANDLE* hItems, vector<string>& itemNames )
{
	string str;
	//string jiange = ".";
	for(int i = 0; i < itemNames.size(); i++)
	{
		//��itemIdƴ������
		//str = string("%1%2%3%4%5%6%7").arg(PLC_HEAD, ".", PLC_NAME, ".", OPC_GROUP_NAME, ".", itemNames[i]);
		/*str += string(OPC_GROUP_NAME);
		str += jiange;
		str += itemNames[i];*/
		str = itemNames[i];
		AddTheItem(pIOPCItemMgt, hItems[i], StringToWchar(str));
	}
}

/*************************************************
1.����hServerItemΪitem������Դ�����ȡitem
2.Item��Value ������varValue�д���
ע��ͨ��pIOPCSyncIOָ�����Read����ʱ����������ȡItem
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

	varValue = pValue[0].vDataValue;//�õ�������

	CoTaskMemFree(pErrors);
	pErrors = NULL;

	CoTaskMemFree(pValue);
	pValue = NULL;

	pIOPCSyncIO->Release();
}

/*******************************************************
������ȡItem�ķ�����������������͵���ˣ�158��ע��Ҳ˵����һ��
1.hItems ΪItem���ָ�루������Ǿ����������
2.itemCount�������Ԫ�صĸ���
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
дItem
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
//����дItem�ķ���
void WriteItems(IOPCItemMgt* pItemMgt, OPCHANDLE* hItems, int itemCount, VARIANT* varValue)
{
	for(int i = 0; i < itemCount; i++)
	{
		WriteItem(pItemMgt, hItems[i], varValue[i]);
	}
}

/**********************************************************
ɾ��һ��Item��IOPCItemMgt::RemoveItems ����ɾ�����Item�Ĺ���
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
ɾ�����Item�ķ��������ﻹ��͵���ˣ�����ֱ��ͨ��IOPCItemMgt::RemoveItems
����ɾ��Item
********************************************************************/
void RemoveItems( IOPCItemMgt* pIOPCItemMgt, OPCHANDLE* hItems, int itemCount)
{
	for(int i = 0; i < itemCount; i++)
	{
		RemoveItem(pIOPCItemMgt, hItems[i]);
	}
}

/***********************************************
�Ƴ�OPC��
************************************************/
void RemoveGroup (IOPCServer* pIOPCServer, OPCHANDLE hServerGroup)
{
	HRESULT hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);
	_ASSERT(!hr);
}



/***********************************************
	�ͷŶ�̬������ڴ�
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
        char* ��wchar_t*��ת������
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
        std stringת����wchar_t*�ķ���
***********************************************/
static wchar_t* StringToWchar(const string& s)  
{  
	const char* p=s.c_str();  
	return CharToWchar(p);  
} 

//�ַ����ָ�����
vector<string> split(string str, string pattern)
{
    string::size_type pos;
    vector<string> result;

    str += pattern;//��չ�ַ����Է������
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

