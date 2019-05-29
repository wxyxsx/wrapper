#include "pch.h"
#include <iostream>
#include <Windows.h>
#include <strsafe.h>
#pragma comment(lib, "User32.lib")

using namespace std;

const char *ACRO_DDESERVER = "AcroViewR19";
const char *ACRO_DDETOPIC = "Control";
const DWORD MAX_TIMEOUT = 3000, STEP_SIZE = 1000;

const char *sDir = "C:\\Users\\wxy\\Desktop\\test";

HDDEDATA CALLBACK DDE_ProcessMessage(UINT uType, UINT uFmt, HCONV hconv, HSZ hsz1, HSZ hsz2,
	HDDEDATA hdata, DWORD dwData1, DWORD dwData2);

int main()
{
	TCHAR ddeCmdBuf[MAX_PATH+10]; 
	UINT retVal;
	DWORD id = 0;
	
	retVal = DdeInitialize(
		&id,
		&DDE_ProcessMessage,
		APPCMD_CLIENTONLY,
		0
	);

	if (retVal != DMLERR_NO_ERROR) {
		cout << "Failed to initialize DDE" << endl;
		return -1;
	}

	DWORD dwResult;
	HCONV hConversation = NULL;
	HSZ hszServerName, hszTopicName;

	hszServerName = DdeCreateStringHandle(id, ACRO_DDESERVER, 0);
	hszTopicName = DdeCreateStringHandle(id, ACRO_DDETOPIC, 0);

	DWORD dwSleep = 0;
	do {
		hConversation = DdeConnect(
			id, 
			hszServerName, 
			hszTopicName, 
			NULL);
		if (hConversation || (dwSleep > MAX_TIMEOUT))
			break;
		Sleep(dwSleep += STEP_SIZE);
	} while (true);

	if (!hConversation) {
		cout << "Could not connect to server." << endl;
		return -1;
	}

	WIN32_FIND_DATA fdFile;
	HANDLE hFind = INVALID_HANDLE_VALUE;
	TCHAR sPath[MAX_PATH];
	
	StringCchCopy(sPath, MAX_PATH, sDir);
	StringCchCat(sPath, MAX_PATH, TEXT("\\*"));

	if ((hFind = FindFirstFile(sPath, &fdFile)) != INVALID_HANDLE_VALUE) {
		do {
			if (fdFile.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
				continue;
			}
			else {
				TCHAR ddeCmdBuf[MAX_PATH + 10];
				snprintf(ddeCmdBuf, MAX_PATH + 10, "[DocOpen(\"%s\\%s\")]", sDir, fdFile.cFileName);
				DdeClientTransaction((unsigned char *)ddeCmdBuf, (DWORD)strlen(ddeCmdBuf), (HCONV)hConversation,
					NULL, (UINT)CF_TEXT, (UINT)XTYP_EXECUTE, (DWORD)1000, &dwResult);
				Sleep(5000);
				snprintf(ddeCmdBuf, MAX_PATH + 10, "[DocClose(\"%s\\%s\")]", sDir, fdFile.cFileName);
				DdeClientTransaction((unsigned char *)ddeCmdBuf, (DWORD)strlen(ddeCmdBuf), (HCONV)hConversation,
					NULL, (UINT)CF_TEXT, (UINT)XTYP_EXECUTE, (DWORD)1000, &dwResult);
				
			}
		} while (FindNextFile(hFind, &fdFile));
	}
	else {
		cout << "Path not found" << endl;
	}

	FindClose(hFind);

	DdeDisconnect(hConversation);
	DdeFreeStringHandle(id, hszServerName);
	DdeFreeStringHandle(id, hszTopicName);
	DdeUninitialize(id);

	return 0;
}

HDDEDATA CALLBACK DDE_ProcessMessage(
	UINT uType,  // Transaction type. 
	UINT uFmt,   // Clipboard data format
	HCONV hconv, // Handle to the conversation
	HSZ hsz1,    // Handle to a string
	HSZ hsz2,    // Handle to a string
	HDDEDATA hdata, // Handle to a global memory object
	DWORD dwData1,  // Transaction-specific data
	DWORD dwData2)  // Transaction-specific data
{
	return NULL;
}