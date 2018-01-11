// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////
#pragma once
#include <initguid.h>

#import <libid:52D76F35-4551-4c56-B53B-A343E42B0AF8> version("2.6") raw_dispinterfaces
#import <libid:74B9FB9B-58D2-4789-9052-ED4EF2ADA75F> version("2.6") raw_dispinterfaces

DEFINE_GUID(LIBID_ProgBarLibU, 0x52D76F35, 0x4551, 0x4c56, 0xB5, 0x3B, 0xA3, 0x43, 0xE4, 0x2B, 0x0A, 0xF8);
DEFINE_GUID(LIBID_ProgBarLibA, 0x74B9FB9B, 0x58D2, 0x4789, 0x90, 0x52, 0xED, 0x4E, 0xF2, 0xAD, 0xA7, 0x5F);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_PROGBARU, CMainDlg, &__uuidof(ProgBarLibU::_IProgressBarEvents), &LIBID_ProgBarLibU, 2, 6>,
    public IDispEventImpl<IDC_PROGBARA, CMainDlg, &__uuidof(ProgBarLibA::_IProgressBarEvents), &LIBID_ProgBarLibA, 2, 6>
{
public:
	enum { IDD = IDD_MAINDLG };

	#define TIMERID_PROGRESSCHANGE 99

	CContainedWindowT<CAxWindow> progBarUContainerWnd;
	CContainedWindowT<CWindow> progBarUWnd;
	CContainedWindowT<CAxWindow> progBarAContainerWnd;
	CContainedWindowT<CWindow> progBarAWnd;

	CMainDlg() :
	    progBarUContainerWnd(this, 1),
	    progBarAContainerWnd(this, 2),
	    progBarUWnd(this, 3),
	    progBarAWnd(this, 4)
	{
	}

	struct Controls
	{
		CEdit logEdit;
		CComPtr<ProgBarLibU::IProgressBar> progBarU;
		CComPtr<ProgBarLibA::IProgressBar> progBarA;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_TIMER, OnTimer)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		ALT_MSG_MAP(4)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 0, ClickProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 1, ContextMenuProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 2, DblClickProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 3, DestroyedControlWindowProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 4, MClickProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 5, MDblClickProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 6, MouseDownProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 7, MouseEnterProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 8, MouseHoverProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 9, MouseLeaveProgbaru)
		//SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 10, MouseMoveProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 11, MouseUpProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 12, OLEDragDropProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 13, OLEDragEnterProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 14, OLEDragLeaveProgbaru)
		//SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 15, OLEDragMouseMoveProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 16, RClickProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 17, RDblClickProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 18, RecreatedControlWindowProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 19, ResizedControlWindowProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 20, XClickProgbaru)
		SINK_ENTRY_EX(IDC_PROGBARU, __uuidof(ProgBarLibU::_IProgressBarEvents), 21, XDblClickProgbaru)

		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 0, ClickProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 1, ContextMenuProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 2, DblClickProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 3, DestroyedControlWindowProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 4, MClickProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 5, MDblClickProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 6, MouseDownProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 7, MouseEnterProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 8, MouseHoverProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 9, MouseLeaveProgbara)
		//SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 10, MouseMoveProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 11, MouseUpProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 12, OLEDragDropProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 13, OLEDragEnterProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 14, OLEDragLeaveProgbara)
		//SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 15, OLEDragMouseMoveProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 16, RClickProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 17, RDblClickProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 18, RecreatedControlWindowProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 19, ResizedControlWindowProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 20, XClickProgbara)
		SINK_ENTRY_EX(IDC_PROGBARA, __uuidof(ProgBarLibA::_IProgressBarEvents), 21, XDblClickProgbara)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_PROGBARA, DLSZ_SIZE_X | DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDC_PROGBARU, DLSZ_SIZE_X | DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDC_EDITLOG, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnTimer(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);

	void __stdcall ClickProgbaru(short button, short shift, long x, long y);
	void __stdcall ContextMenuProgbaru(short button, short shift, long x, long y);
	void __stdcall DblClickProgbaru(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowProgbaru(long hWnd);
	void __stdcall MClickProgbaru(short button, short shift, long x, long y);
	void __stdcall MDblClickProgbaru(short button, short shift, long x, long y);
	void __stdcall MouseDownProgbaru(short button, short shift, long x, long y);
	void __stdcall MouseEnterProgbaru(short button, short shift, long x, long y);
	void __stdcall MouseHoverProgbaru(short button, short shift, long x, long y);
	void __stdcall MouseLeaveProgbaru(short button, short shift, long x, long y);
	void __stdcall MouseMoveProgbaru(short button, short shift, long x, long y);
	void __stdcall MouseUpProgbaru(short button, short shift, long x, long y);
	void __stdcall OLEDragDropProgbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterProgbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveProgbaru(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveProgbaru(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickProgbaru(short button, short shift, long x, long y);
	void __stdcall RDblClickProgbaru(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowProgbaru(long hWnd);
	void __stdcall ResizedControlWindowProgbaru();
	void __stdcall XClickProgbaru(short button, short shift, long x, long y);
	void __stdcall XDblClickProgbaru(short button, short shift, long x, long y);

	void __stdcall ClickProgbara(short button, short shift, long x, long y);
	void __stdcall ContextMenuProgbara(short button, short shift, long x, long y);
	void __stdcall DblClickProgbara(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowProgbara(long hWnd);
	void __stdcall MClickProgbara(short button, short shift, long x, long y);
	void __stdcall MDblClickProgbara(short button, short shift, long x, long y);
	void __stdcall MouseDownProgbara(short button, short shift, long x, long y);
	void __stdcall MouseEnterProgbara(short button, short shift, long x, long y);
	void __stdcall MouseHoverProgbara(short button, short shift, long x, long y);
	void __stdcall MouseLeaveProgbara(short button, short shift, long x, long y);
	void __stdcall MouseMoveProgbara(short button, short shift, long x, long y);
	void __stdcall MouseUpProgbara(short button, short shift, long x, long y);
	void __stdcall OLEDragDropProgbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterProgbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveProgbara(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveProgbara(LPDISPATCH data, long* effect, short button, short shift, long x, long y);
	void __stdcall RClickProgbara(short button, short shift, long x, long y);
	void __stdcall RDblClickProgbara(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowProgbara(long hWnd);
	void __stdcall ResizedControlWindowProgbara();
	void __stdcall XClickProgbara(short button, short shift, long x, long y);
	void __stdcall XDblClickProgbara(short button, short shift, long x, long y);
};
