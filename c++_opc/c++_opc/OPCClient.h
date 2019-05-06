#ifndef _OPC_CLIENT_H
#define _OPC_CLIENT_H

//#include <QString>
//#include <QStringList>
using namespace std;

IOPCServer* InstantiateServer(wchar_t ServerName[]);
void AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt* &pIOPCItemMgt, 
				 OPCHANDLE& hServerGroup, wchar_t groupName[]);
void AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem, wchar_t itemId[]);
void ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue);
void ReadItems(IOPCItemMgt* pItemMgt, OPCHANDLE* hItems, int itemCount, VARIANT& varValue);
void WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE* hServerItem, VARIANT& varValue);
void WriteItems(IOPCItemMgt* pItemMgt, OPCHANDLE* hItems, int itemCount, VARIANT* varValue);

void RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE hServerItem);
void RemoveItems(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE* hItems, int itemCount);
void RemoveGroup (IOPCServer* pIOPCServer, OPCHANDLE hServerGroup);
//void AddItemsToGroup(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE* hItems, QStringList& itemNames);
void AddItemsToGroup( IOPCItemMgt* pIOPCItemMgt, OPCHANDLE* hItems, vector<string>& itemNames );

static void Release();
static wchar_t* CharToWchar(const char* c);
static wchar_t* StringToWchar(const string& s);
vector<string> split(string str, string pattern);


#endif // _OPC_CLIENT_H 