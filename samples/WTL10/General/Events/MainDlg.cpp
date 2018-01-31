// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

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
		IDispEventImpl<IDC_PROGBARU, CMainDlg, &__uuidof(ProgBarLibU::_IProgressBarEvents), &LIBID_ProgBarLibU, 2, 6>::DispEventUnadvise(controls.progBarU);
		controls.progBarU.Release();
	}
	if(controls.progBarA) {
		IDispEventImpl<IDC_PROGBARA, CMainDlg, &__uuidof(ProgBarLibA::_IProgressBarEvents), &LIBID_ProgBarLibA, 2, 6>::DispEventUnadvise(controls.progBarA);
		controls.progBarA.Release();
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
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	controls.logEdit = GetDlgItem(IDC_EDITLOG);

	progBarUContainerWnd.SubclassWindow(GetDlgItem(IDC_PROGBARU));
	progBarUContainerWnd.QueryControl(__uuidof(ProgBarLibU::IProgressBar), reinterpret_cast<LPVOID*>(&controls.progBarU));
	if(controls.progBarU) {
		IDispEventImpl<IDC_PROGBARU, CMainDlg, &__uuidof(ProgBarLibU::_IProgressBarEvents), &LIBID_ProgBarLibU, 2, 6>::DispEventAdvise(controls.progBarU);
		progBarUWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.progBarU->GethWnd())));
	}

	progBarAContainerWnd.SubclassWindow(GetDlgItem(IDC_PROGBARA));
	progBarAContainerWnd.QueryControl(__uuidof(ProgBarLibA::IProgressBar), reinterpret_cast<LPVOID*>(&controls.progBarA));
	if(controls.progBarA) {
		IDispEventImpl<IDC_PROGBARA, CMainDlg, &__uuidof(ProgBarLibA::_IProgressBarEvents), &LIBID_ProgBarLibA, 2, 6>::DispEventAdvise(controls.progBarA);
		progBarAWnd.SubclassWindow(static_cast<HWND>(LongToHandle(controls.progBarA->GethWnd())));
	}

	SetTimer(TIMERID_PROGRESSCHANGE, 750);
	return TRUE;
}

LRESULT CMainDlg::OnTimer(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled)
{
	switch(wParam) {
		case TIMERID_PROGRESSCHANGE:
			if(controls.progBarU) {
				controls.progBarU->ChangeCurrentValue(0);
				if(controls.progBarU->GetCurrentValue() == controls.progBarU->GetMaximum()) {
					controls.progBarU->PutStepWidth(-10);
				} else if(controls.progBarU->GetCurrentValue() == controls.progBarU->GetMinimum()) {
					controls.progBarU->PutStepWidth(10);
				}
			}
			if(controls.progBarA) {
				controls.progBarA->ChangeCurrentValue(0);
				if(controls.progBarA->GetCurrentValue() == controls.progBarA->GetMaximum()) {
					controls.progBarA->PutStepWidth(-10);
				} else if(controls.progBarA->GetCurrentValue() == controls.progBarA->GetMinimum()) {
					controls.progBarA->PutStepWidth(10);
				}
			}
			break;
		default:
			bHandled = FALSE;
			break;
	}

	return 0;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.progBarU) {
		controls.progBarU->About();
	}
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void __stdcall CMainDlg::ClickProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowProgbaru(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropProgbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ProgBarLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ProgBarU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ProgBarU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.progBarU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterProgbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ProgBarLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ProgBarU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ProgBarU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveProgbaru(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ProgBarLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ProgBarU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ProgBarU_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveProgbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ProgBarLibU::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ProgBarU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ProgBarU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowProgbaru(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowProgbaru()
{
	AddLogEntry(CAtlString(TEXT("ProgBarU_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickProgbaru(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarU_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}


void __stdcall CMainDlg::ClickProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_Click: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_ContextMenu: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_DblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowProgbara(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_MClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_MDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_MouseDown: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_MouseEnter: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_MouseHover: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_MouseLeave: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_MouseMove: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_MouseUp: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropProgbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ProgBarLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ProgBarA_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ProgBarA_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.progBarA->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterProgbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ProgBarLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ProgBarA_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ProgBarA_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveProgbara(LPDISPATCH data, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ProgBarLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ProgBarA_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ProgBarA_OLEDragLeave: data=NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveProgbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y)
{
	CAtlString str;
	CComQIPtr<ProgBarLibA::IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("ProgBarA_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("ProgBarA_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, button=%i, shift=%i, x=%i, y=%i"), *effect, button, shift, x, y);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_RClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_RDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowProgbara(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::ResizedControlWindowProgbara()
{
	AddLogEntry(CAtlString(TEXT("ProgBarA_ResizedControlWindow")));
}

void __stdcall CMainDlg::XClickProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_XClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickProgbara(short button, short shift, long x, long y)
{
	CAtlString str;
	str.Format(TEXT("ProgBarA_XDblClick: button=%i, shift=%i, x=%i, y=%i"), button, shift, x, y);

	AddLogEntry(str);
}
