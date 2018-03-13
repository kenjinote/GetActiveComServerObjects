#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>

TCHAR szClassName[] = TEXT("Window");

void GetActiveComServerObjects(HWND hList)
{
	SendMessage(hList, LB_RESETCONTENT, 0, 0);
	DWORD dwSubKeys = 0;
	DWORD retCode = RegQueryInfoKey(HKEY_CLASSES_ROOT, 0, 0, 0, &dwSubKeys, 0, 0, 0, 0, 0, 0, 0);
	if (retCode == ERROR_SUCCESS && dwSubKeys > 0)
	{
		for (DWORD i = 0; i < dwSubKeys; ++i)
		{
			TCHAR szKeyName[256];
			DWORD cbName = _countof(szKeyName);
			retCode = RegEnumKeyEx(HKEY_CLASSES_ROOT, i, szKeyName, &cbName, 0, 0, 0, 0);
			if (retCode == ERROR_SUCCESS)
			{
				TCHAR szKeyPath[1024];
				lstrcpy(szKeyPath, szKeyName);
				lstrcat(szKeyPath, TEXT("\\"));
				lstrcat(szKeyPath, TEXT("CLSID"));
				HKEY hKey;
				if (RegOpenKeyEx(HKEY_CLASSES_ROOT, szKeyPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
				{
					DWORD dwType;
					TCHAR lpData[256];
					DWORD dwDataSize = _countof(lpData);
					retCode = RegQueryValueEx(hKey, NULL, 0, &dwType, (LPBYTE)lpData, &dwDataSize);
					if (retCode == ERROR_SUCCESS)
					{
						CLSID clsid;
						CLSIDFromString(lpData, &clsid);
						LPUNKNOWN pUnknown = NULL;
						if (S_OK == GetActiveObject(clsid, NULL, &pUnknown))
						{
							SendMessage(hList, LB_ADDSTRING, 0, (LPARAM)szKeyName);
							pUnknown->Release();
						}
					}
					RegCloseKey(hKey);
				}
			}
		}
	}
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hButton;
	static HWND hList;
	switch (msg)
	{
	case WM_CREATE:
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("更新"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hList = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("LISTBOX"), 0, WS_VISIBLE | WS_CHILD | WS_VSCROLL | LBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		PostMessage(hWnd, WM_COMMAND, IDOK, 0);
		break;
	case WM_SIZE:
		MoveWindow(hButton, 10, 10, 256, 32, TRUE);
		MoveWindow(hList, 10, 50, LOWORD(lParam) - 20, HIWORD(lParam) - 60, TRUE);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			GetActiveComServerObjects(hList);
		}
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("現在アクティブなCOMサーバー一覧"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
