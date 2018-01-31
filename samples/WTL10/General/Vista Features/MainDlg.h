// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:52D76F35-4551-4c56-B53B-A343E42B0AF8> version("2.6") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_PROGBARU, CMainDlg, &__uuidof(_IProgressBarEvents), &LIBID_ProgBarLibU, 2, 6>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> progBarUWnd;

	CMainDlg() : progBarUWnd(this, 1)
	{
	}

	struct Controls
	{
		CComPtr<IProgressBar> progBarU;
		CButton stateButton[3];
		CButton smoothReverseButton;
		CEdit currentValueEdit;
		CButton applyButton;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		COMMAND_ID_HANDLER(IDC_APPLY, OnApply)
		COMMAND_HANDLER(IDC_CURRENTVALUE, EN_CHANGE, OnCurrentValueChange)
		COMMAND_ID_HANDLER(IDC_SMOOTHREVERSE, OnSmoothReverse)
		COMMAND_ID_HANDLER(IDC_STATENORMAL, OnStateChanged)
		COMMAND_ID_HANDLER(IDC_STATEPAUSED, OnStateChanged)
		COMMAND_ID_HANDLER(IDC_STATEFAILED, OnStateChanged)

		ALT_MSG_MAP(1)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
	END_SINK_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnApply(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCurrentValueChange(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSmoothReverse(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnStateChanged(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void CloseDialog(int nVal);
};
