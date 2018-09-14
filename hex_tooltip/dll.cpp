
/*
	Author: david zimmer 
	email:  dzzie@yahoo.com
	site:   http://sandsprite.com

	inject this into the vb6 ide process and it will hook the tooltips.
	if it detects a numeric tooltip value is about to display it will
	modify the tool tip text and display the original text plus its value in hex.

	portions of this code subject to license below
*/

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

#define _WIN32_WINNT 0x0401  //for IsDebuggerPresent 
#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stdarg.h>
#pragma warning(disable:4996)
void InstallHooks(void);

#include "NtHookEngine.h"
#include "main.h"          //contains a bunch of library functions in it too..

bool Installed =false;
bool SubclassActive = false;
char* newText = NULL;
WNDPROC OldWndProc = NULL;
int WindowWidth = 0;
HWND hookedHWND = 0;   //tooltips use the same HWND for every run and every difference code editor window..we only have to hook once..
HFONT hFont;
const bool dbgMsg = false;
#define ERROR_NO_KEY      0x11223344
char my_regKey[200] = "Software\\VB and VBA Program Settings\\FastBuild\\Settings";


void Closing(void){ msg("***** Injected Process Terminated *****"); exit(0);}
	
extern "C" __declspec (dllexport) int NullSub(void){ return 1;} //so we have an export to hardcode add to pe table if we want.

BOOL APIENTRY DllMain( HANDLE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
    if(!Installed){
		 Installed=true;
		 InstallHooks();
		 atexit(Closing);

		 hFont = CreateFont(14,0,0,0,FW_DONTCARE,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_OUTLINE_PRECIS,
							CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, DEFAULT_PITCH, TEXT("Arial")
				 ); 

	}

	return TRUE;
}


int getStringWidth(char *text, HFONT font) {
    HDC dc = GetDC(NULL);
    SelectObject(dc, font);

    RECT rect = { 0, 0, 0, 0 };
    DrawText(dc, text, strlen(text), &rect, DT_CALCRECT | DT_NOPREFIX | DT_SINGLELINE);
    int textWidth = abs(rect.right - rect.left);

    ReleaseDC(NULL,dc);
    return textWidth;
}

LRESULT APIENTRY NewWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	PAINTSTRUCT ps;
	HDC hdc;
    RECT rc;

	if(SubclassActive && newText != NULL){
		int txtLen = strlen(newText);
		switch (uMsg)
		{
			case WM_PAINT:

				 hdc = BeginPaint(hwnd,&ps);

				 SetBkColor(hdc, RGB(255,255,0xe1));
				 SetTextColor(hdc, RGB(0,0,0));
				 SelectObject(hdc,hFont);
				 TextOut(hdc,0,0, newText, txtLen);

 				 GetWindowRect(hwnd,&rc);
				 HBRUSH brush = CreateSolidBrush(RGB(255,255,0xe1));
				 if(brush!=NULL){
					 FillRect(hdc, &rc, brush);
					 DeleteObject(brush);
				 }

				 EndPaint(hwnd,&ps);  
				
				 return 0;
		}
	}
	return CallWindowProc(OldWndProc, hwnd, uMsg, wParam, lParam);
}

int AppendHexIfOk(HWND h, char* curCaption){
	
	char c;
	int equalPos = 0;

	if(curCaption==NULL) return 0;

	int curLen = strlen(curCaption);
	if(curLen==0) return 0;

	//first scan and see if its a valid target..
	for(int i=0; i < curLen; i++){
		c = curCaption[i];
		if(c == '='){ equalPos = i; continue; /*jump to next i*/}
		if(equalPos > 0){
			if(c==' ' || isdigit(c)){ /*ok*/ } else return 0; //only accept spaces and numbers for this mod..
		}
	}

	if(equalPos == 0) return 0;
	
	//DebugBreak();

	//so all tooltips for the debugger use the same tooltip hwnd..once we set it once were good..what about different editor windows..
	if(OldWndProc == NULL){//SetWindowText does not work we must owner draw or something...
		hookedHWND = h;
		OldWndProc = (WNDPROC)SetWindowLongPtr (h, GWLP_WNDPROC, (LONG_PTR)NewWndProc);
	}

	if(newText != NULL)	free(newText);
	newText = (char*)malloc(curLen + 100);
	int numericVal = atoi((char*)(curCaption+equalPos+1));
	sprintf(newText, "  %s   [0x%X] ", curCaption, numericVal);
	if(dbgMsg) LogAPI("Modified ToolTip(%x) newText=%s", h, newText);
    
	return 1;
}

char* windowText(HWND h){
	
	char *caption = NULL;
	int cLen = GetWindowTextLengthA(h);
	if(cLen == 0) goto err;

	cLen+=20;
	caption = (char*)malloc(cLen);
	if(caption==0)  goto err;

	if(GetWindowText(h, caption, cLen) == 0){
		 free(caption);
		 goto err;
	}

	 return caption;
err: return strdup("");
}

void resizeToolTip(HWND hwnd){
	 RECT rc;
	 if(newText == NULL) return;
	 GetWindowRect(hwnd, &rc);
	 int width = getStringWidth(newText, hFont); //rc.left - rc.right + 500;  
	 int height = rc.bottom - rc.top;
	 Real_SetWindowPos(hwnd, NULL,  rc.left, rc.top, width, height, SWP_SHOWWINDOW);
	 //SendMessage(hWnd, TTM_ADJUSTRECT, TRUE, (LPARAM)&rc); 			  
}

int ReadRegInt(char* baseKey, char* name){

	 char tmp[20] = {0};
     unsigned long l = sizeof(tmp);
	 HKEY h;
	 
	 int rv = RegOpenKeyEx(HKEY_CURRENT_USER, baseKey, 0, KEY_READ, &h);
	 rv = RegQueryValueExA(h, name, 0,0, (unsigned char*)tmp, &l);
	 RegCloseKey(h);

	 if(rv != ERROR_SUCCESS) return ERROR_NO_KEY;
	 return atoi(tmp);
}

BOOL __stdcall My_SetWindowPos(HWND hWnd, HWND hWndInsertAfter, int  X, int  Y, int  cx, int  cy, UINT uFlags){

	BOOL rv;
	char buf[500] = {0};
	
	SubclassActive = false;
	int sz = GetClassName(hWnd, buf, sizeof(buf));
	
	if(sz!=0 && strcmp(buf,"tooltips_class32") == 0 && (uFlags & SWP_SHOWWINDOW) == SWP_SHOWWINDOW){
		
		int v = ReadRegInt(my_regKey, "DisplayAsHex");

		if(v==0){
			if(dbgMsg) LogAPI("%x  SetWindowPos Hook Disabled h=%x   flags=%x", CalledFrom(), hWnd, uFlags);
		}else{ 
			if(v == ERROR_NO_KEY && dbgMsg) LogAPI("SetWindowPos Hook no reg key set continuing...");

			char *caption = windowText(hWnd); //always returns a malloced buf to free
			
			if(AppendHexIfOk(hWnd,caption)==1){
				SubclassActive = true;
				if(dbgMsg) LogAPI("%x  Modified ToolTip.SetWindowPos h=%x   flags=%x   org=%s (%d,%d,%d,%d)", CalledFrom(), hWnd, uFlags, caption,X,Y,cx,cy);
				rv = Real_SetWindowPos(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
				resizeToolTip(hWnd);
				free(caption);
				return rv;
			}
			else{
				if(dbgMsg) LogAPI("%x  No modify ToolTip.SetWindowPos h=%x   flags=%x   %s", CalledFrom(), hWnd, uFlags, caption);
				free(caption);
			}
		}

		
	}else{
		if(dbgMsg) LogAPI("%x  SetWindowPos h=%x flags=%x  class= %s (x=%d,y=%d,cx=%d,cy=%d)", CalledFrom(), hWnd, uFlags, buf, X,Y,cx,cy);
	}

	
	return Real_SetWindowPos(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);

}




//_______________________________________________ install hooks fx 

bool InstallHook( void* real, void* hook, int* thunk, char* name, enum hookType ht){
	if( HookFunction((ULONG_PTR) real, (ULONG_PTR)hook, name, ht) ){ 
		*thunk = (int)GetOriginalFunction((ULONG_PTR) hook);
		return true;
	}
	return false;
}

//before it was depending on getting the real address from import table, now its explicitly retreived from specified dll
HMODULE hKernelBase = 0;
char* curDLL = NULL; 

void DoHook(void* hook, int* thunk, char* name){

	void *lpReal = 0;
	HMODULE dllBase = 0;

	if(curDLL==NULL){
		LogAPI("Install %s hook failed...curDll is null..\r\n", name);
		return;
	}

	if(strstr(curDLL, "kernel") > 0){
		if(hKernelBase != 0){//its Vista+, see if the export exists there. if its in both we want kernelBase version instead..
			//if(Real_GetProcAddress == NULL){
					lpReal = (void*)GetProcAddress(hKernelBase, name); //k32 is just a forwarder which we cant hook...
			//}else{
			//		lpReal = (void*)Real_GetProcAddress(hKernelBase, name); 
			//} 
		}
	}
	
	if(lpReal==0){ //it wasnt kernelxx, or not vista+ or not in kernelbase.dll
		dllBase = GetModuleHandle(curDLL);
		if(dllBase==0) dllBase = LoadLibrary(curDLL);
		
		if(dllBase==0){
			LogAPI("Install %s hook failed...%s could not be loaded..\r\n", name, curDLL);
			return;
		}

		//if(Real_GetProcAddress == NULL){
				lpReal = (void*)GetProcAddress(dllBase, name); //k32 is just a forwarder which we cant hook...
		//}else{
		//		lpReal = (void*)Real_GetProcAddress(dllBase, name); 
		//} 
	}

	if(lpReal==0){ 
		LogAPI("Install %s hook failed...could not find address in %s..\r\n", name, curDLL);
		return;
	}

	if(!InstallHook( lpReal, hook, thunk, name, ht_auto ) ){
		LogAPI("Install %s hook failed...\r\nError: %s\r\n", name, GetHookError());
	}

	 
}


//Macro wrapper to build DoHook() call
#define ADDHOOK(name) DoHook( My_##name, (int*)&Real_##name, #name );
	
void HookEngineDebugMessage(char* msg){
	LogAPI("Debug> %s", msg);
}

void InstallHooks(void)
{

	logLevel = 0;
	debugMsgHandler = HookEngineDebugMessage;

	msg("***** Installing Hooks *****");	
	
	//DO NOT HOOK GetProcAddress or GetModuleHandle we use them below (not in hook engine)..
	hKernelBase = GetModuleHandle("kernelbase.dll");

	//before you disable any of these hooks MAKE SURE to grep the source to make sure the 
	//Real_ versions arent in use in this lib.. or its boom boom time. (not in the good way)
	
	//note 2: if you set a regular breakpoint on a hooked windows api, before it is hooked, 
	//it will fuckup the hook engine because the int3 will be copied to the thunk. 

	curDLL = "User32.dll";
	ADDHOOK(SetWindowPos)
	
}
