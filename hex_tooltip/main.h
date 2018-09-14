/*
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA
*/

#include <intrin.h>
#include <psapi.h>
#include <Shlobj.h>
#include <direct.h>

#pragma comment(lib, "psapi.lib")
#pragma comment(lib, "Shell32.lib")

BOOL (__stdcall *Real_SetWindowPos)(HWND hWnd, HWND hWndInsertAfter, int  X, int  Y, int  cx, int  cy, UINT uFlags) = NULL;

void msg(char);
void LogAPI(const char*, ...);

bool Warned=false;
HWND hServer=0;
int DumpAt=0;
char *dllPath = 0; //fullpath to api_log.dll
char *wpmPath = 0; //WriteProcessMemory Dump path

extern int myPID;

bool FolderExists(char* folder)
{
	DWORD rv = GetFileAttributes(folder);
	if( rv == INVALID_FILE_ATTRIBUTES) return false;
	if( !(rv & FILE_ATTRIBUTE_DIRECTORY) ) return false;
	return true;
}


char* FileNameFromPath(char* path){
	if(path==NULL || strlen(path)==0) return strdup("");
	unsigned int x = strlen(path);
	while(x > 0){
		if( path[x-1] == '\\') break;
		x--;
	}
	int sz = strlen(path) - x;
	char* tmp = (char*)malloc(sz+2);
	memset(tmp,0,sz+2);
	for(int i=0; i < sz; i++){
		tmp[i] = path[x+i];
	}
	return tmp;
}


void FindVBWindow(){
	char *vbIDEClassName = "ThunderFormDC" ;
	char *vbEXEClassName = "ThunderRT6FormDC" ;
	char *vbWindowCaption = "ApiLogger" ;

	hServer = FindWindowA( vbIDEClassName, vbWindowCaption );
	if(hServer==0) hServer = FindWindowA( vbEXEClassName, vbWindowCaption );

	if(hServer==0){
		if(!Warned){
			//MessageBox(0,"Could not find msg window","",0);
			Warned=true;
		}
	}
	else{
		if(!Warned){
			//first time we are being called we could do stuff here...
			Warned=true;

		}
	}	

} 

char msgbuf[0x1001];

int msg(char *Buffer){
  
  if(!IsWindow(hServer)) hServer=0;
  if(hServer==0) FindVBWindow();
  int myPID = GetCurrentProcessId();

  COPYDATASTRUCT cpStructData;
  memset(&cpStructData,0, sizeof(struct tagCOPYDATASTRUCT )) ;
  
  _snprintf(msgbuf, 0x1000, "%x,%x,%s", myPID, GetCurrentThreadId(), Buffer);

  cpStructData.dwData = 3;
  cpStructData.cbData = strlen(msgbuf) ;
  cpStructData.lpData = (void*)msgbuf;
  
  int ret = SendMessage(hServer, WM_COPYDATA, 0,(LPARAM)&cpStructData);

  //if ret == 0x then do something special like reconfig ?

  return ret;

} 

void LogAPI(const char *format, ...)
{
	DWORD dwErr = GetLastError();
		
	if(format){
		char buf[1024]; 
		va_list args; 
		va_start(args,format); 
		try{
 			 _vsnprintf(buf,1024,format,args);
			 msg(buf);
		}
		catch(...){}
	}

	SetLastError(dwErr);
}

#define	CalledFrom() (int)_ReturnAddress()

/*__declspec(naked) int CalledFrom(){ 
	
	_asm{
			 mov eax, [ebp+4]  //return address of parent function (were nekkid)
			 ret
	}
	
}*/

 

