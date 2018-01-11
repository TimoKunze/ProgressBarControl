// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include <olectl.h>
#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(controls.progBarU) {
		DispEventUnadvise(controls.progBarU);
		controls.progBarU.Release();
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.stateButton[0] = GetDlgItem(IDC_STATENORMAL);
	controls.stateButton[0].SetCheck(BST_CHECKED);
	controls.stateButton[1] = GetDlgItem(IDC_STATEPAUSED);
	controls.stateButton[2] = GetDlgItem(IDC_STATEFAILED);
	controls.smoothReverseButton = GetDlgItem(IDC_SMOOTHREVERSE);
	controls.smoothReverseButton.SetCheck(BST_CHECKED);
	controls.applyButton = GetDlgItem(IDC_APPLY);
	controls.currentValueEdit = GetDlgItem(IDC_CURRENTVALUE);
	controls.currentValueEdit.SetWindowText(TEXT("0"));
	controls.applyButton.EnableWindow(FALSE);

	progBarUWnd.SubclassWindow(GetDlgItem(IDC_PROGBARU));
	progBarUWnd.QueryControl(__uuidof(IProgressBar), reinterpret_cast<LPVOID*>(&controls.progBarU));
	if(controls.progBarU) {
		DispEventAdvise(controls.progBarU);
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.progBarU) {
		controls.progBarU->About();
	}
	return 0;
}

LRESULT CMainDlg::OnApply(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.progBarU) {
		CAtlString str;
		controls.currentValueEdit.GetWindowText(str);
		long value = _ttol(str);
		controls.progBarU->PutCurrentValue(value);

		TCHAR* pBuffer = new TCHAR[70];
		ZeroMemory(pBuffer, 70 * sizeof(TCHAR));
		_ltot_s(value, pBuffer, 70, 10);
		controls.currentValueEdit.SetWindowText(pBuffer);

		controls.applyButton.EnableWindow(FALSE);
	}
	return 0;
}

LRESULT CMainDlg::OnCurrentValueChange(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	controls.applyButton.EnableWindow(TRUE);
	return 0;
}

LRESULT CMainDlg::OnSmoothReverse(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.progBarU) {
		controls.progBarU->PutSmoothReverse((controls.smoothReverseButton.GetCheck() == BST_CHECKED) ? VARIANT_TRUE : VARIANT_FALSE);
	}
	return 0;
}

LRESULT CMainDlg::OnStateChanged(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.progBarU) {
		if(controls.stateButton[0].GetCheck() == BST_CHECKED) {
			controls.progBarU->PutProgressState(psNormal);
		} else if(controls.stateButton[1].GetCheck() == BST_CHECKED) {
			controls.progBarU->PutProgressState(psPaused);
		} else if(controls.stateButton[2].GetCheck() == BST_CHECKED) {
			controls.progBarU->PutProgressState(psFailed);
		}
	}
	return 0;
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}
