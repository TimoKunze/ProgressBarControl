// ProgressBar.cpp: Superclasses msctls_progress32.

#include "stdafx.h"
#include "ProgressBar.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
FONTDESC ProgressBar::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


ProgressBar::ProgressBar()
{
	properties.font.InitializePropertyWatcher(this, DISPID_PROGBAR_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_PROGBAR_MOUSEICON);

	SIZEL size = {100, 20};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP ProgressBar::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IProgressBar, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ProgressBar::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<ProgressBar>::Load(pPropertyBag, pErrorLog);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);

		CComBSTR bstr;
		bstr = L"%1 %%";
		hr = pPropertyBag->Read(OLESTR("Text"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_Text(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP ProgressBar::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<ProgressBar>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.text.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("Text"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP ProgressBar::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 21 VT_I4 properties...
	pSize->QuadPart += 21 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 11 VT_BOOL properties...
	pSize->QuadPart += 11 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 1 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.text.ByteLength() + sizeof(OLECHAR);

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP ProgressBar::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	// NOTE: We neither support legacy streams nor streams with a version < 0x0100.

	HRESULT hr = S_OK;
	LONG signature = 0;
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x07070707/*4x AppID*/) {
		// might be a legacy stream, that starts with the ATL version
		if(signature == 0x0700 || signature == 0x0710 || signature == 0x0800 || signature == 0x0900) {
			version = 0x0099;
		} else {
			return E_FAIL;
		}
	}
	//LONG version = 0;
	if(version != 0x0099) {
		if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
			return hr;
		}
		if(version > 0x0103) {
			return E_FAIL;
		}
	}

	DWORD atlVersion;
	if(version == 0x0099) {
		atlVersion = 0x0900;
	} else {
		if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
			return hr;
		}
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(version != 0x0100) {
		if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
			return hr;
		}
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = NULL;
	if(version == 0x0100) {
		pfnReadVariantFromStream = ReadVariantFromStream_Legacy;
	} else {
		pfnReadVariantFromStream = ReadVariantFromStream;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL activateMarquee = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR barColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BarStyleConstants barStyle = static_cast<BarStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG currentValue = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG marqueeStepDuration = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maximum = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG minimum = propertyValue.lVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OrientationConstants orientation = static_cast<OrientationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ProgressStateConstants progressState = static_cast<ProgressStateConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL rightToLeftLayout = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL smoothReverse = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG stepWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;

	VARIANT_BOOL detectDoubleClicks = VARIANT_TRUE;
	VARIANT_BOOL displayProgressInTaskBar = VARIANT_FALSE;
	VARIANT_BOOL displayText = VARIANT_FALSE;
	CComPtr<IFontDisp> pFont = NULL;
	OLE_COLOR foreColor = RGB(255, 255, 255);
	HAlignmentConstants hAlignment = halCenter;
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(0);
	CComBSTR text;
	OLE_COLOR textShadowColor = RGB(48, 48, 48);
	OLE_XSIZE_PIXELS textShadowOffsetX = 1;
	OLE_YSIZE_PIXELS textShadowOffsetY = 1;
	VARIANT_BOOL useSystemFont = VARIANT_TRUE;
	if(version >= 0x0102) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		detectDoubleClicks = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		displayText = propertyValue.boolVal;

		if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
			return hr;
		}
		if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
			if(hr != REGDB_E_CLASSNOTREG) {
				return S_OK;
			}
		}

		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		foreColor = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		hAlignment = static_cast<HAlignmentConstants>(propertyValue.lVal);
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
		if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
			return hr;
		}
		if(FAILED(hr = text.ReadFromStream(pStream))) {
			return hr;
		}
		if(!text) {
			text = L"";
		}
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		textShadowColor = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		textShadowOffsetX = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		textShadowOffsetY = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		useSystemFont = propertyValue.boolVal;

		if(version >= 0x0103) {
			if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
				return hr;
			}
			displayProgressInTaskBar = propertyValue.boolVal;
		}
	} else {
		text = L"%1 %%";
	}


	hr = put_ActivateMarquee(activateMarquee);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BarColor(barColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BarStyle(barStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CurrentValue(currentValue);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DetectDoubleClicks(detectDoubleClicks);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayProgressInTaskBar(displayProgressInTaskBar);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayText(displayText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HAlignment(hAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MarqueeStepDuration(marqueeStepDuration);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Maximum(maximum);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Minimum(minimum);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Orientation(orientation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProgressState(progressState);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	if(version >= 0x0102) {
		hr = put_RightToLeft(rightToLeft);
		if(FAILED(hr)) {
			return hr;
		}
	} else {
		hr = put_RightToLeftLayout(rightToLeftLayout);
		if(FAILED(hr)) {
			return hr;
		}
	}
	hr = put_SmoothReverse(smoothReverse);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_StepWidth(stepWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TextShadowColor(textShadowColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TextShadowOffsetX(textShadowOffsetX);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_TextShadowOffsetY(textShadowOffsetY);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ProgressBar::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x07070707/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0103;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.activateMarquee);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.barColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.barStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.currentValue;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.marqueeStepDuration;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maximum;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.minimum;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.orientation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.progressState;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rightToLeft & rtlLayout);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.smoothReverse);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.stepWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0102 starts here
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayText);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.text.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.lVal = properties.textShadowColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.textShadowOffsetX;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.textShadowOffsetY;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	// version 0x0103 starts here
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayProgressInTaskBar);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND ProgressBar::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_PROGRESS_CLASS;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<ProgressBar>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT ProgressBar::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("ProgressBar ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void ProgressBar::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP ProgressBar::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<ProgressBar>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	return hr;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP ProgressBar::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP ProgressBar::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP ProgressBar::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_PROGBAR_APPEARANCE:
		case DISPID_PROGBAR_BARSTYLE:
		case DISPID_PROGBAR_BORDERSTYLE:
		case DISPID_PROGBAR_DISPLAYPROGRESSINTASKBAR:
		case DISPID_PROGBAR_DISPLAYTEXT:
		case DISPID_PROGBAR_HALIGNMENT:
		case DISPID_PROGBAR_MOUSEICON:
		case DISPID_PROGBAR_MOUSEPOINTER:
		case DISPID_PROGBAR_ORIENTATION:
		case DISPID_PROGBAR_PROGRESSSTATE:
		case DISPID_PROGBAR_RIGHTTOLEFTLAYOUT:
		case DISPID_PROGBAR_TEXTSHADOWOFFSETX:
		case DISPID_PROGBAR_TEXTSHADOWOFFSETY:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_PROGBAR_ACTIVATEMARQUEE:
		case DISPID_PROGBAR_DETECTDOUBLECLICKS:
		case DISPID_PROGBAR_DISABLEDEVENTS:
		case DISPID_PROGBAR_DONTREDRAW:
		case DISPID_PROGBAR_HOVERTIME:
		case DISPID_PROGBAR_MARQUEESTEPDURATION:
		case DISPID_PROGBAR_RIGHTTOLEFT:
		case DISPID_PROGBAR_SMOOTHREVERSE:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_PROGBAR_BACKCOLOR:
		case DISPID_PROGBAR_BARCOLOR:
		case DISPID_PROGBAR_FORECOLOR:
		case DISPID_PROGBAR_TEXTSHADOWCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_PROGBAR_APPID:
		case DISPID_PROGBAR_APPNAME:
		case DISPID_PROGBAR_APPSHORTNAME:
		case DISPID_PROGBAR_BUILD:
		case DISPID_PROGBAR_CHARSET:
		case DISPID_PROGBAR_HAPPLICATIONWINDOW:
		case DISPID_PROGBAR_HWND:
		case DISPID_PROGBAR_ISRELEASE:
		case DISPID_PROGBAR_PROGRAMMER:
		case DISPID_PROGBAR_TESTER:
		case DISPID_PROGBAR_TEXT:
		case DISPID_PROGBAR_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_PROGBAR_REGISTERFOROLEDRAGDROP:
		case DISPID_PROGBAR_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_PROGBAR_FONT:
		case DISPID_PROGBAR_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_PROGBAR_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
		case DISPID_PROGBAR_CURRENTVALUE:
		case DISPID_PROGBAR_MAXIMUM:
		case DISPID_PROGBAR_MINIMUM:
		case DISPID_PROGBAR_STEPWIDTH:
			*pCategory = PROPCAT_Scale;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString ProgressBar::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ProgressBar::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ProgressBar::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ProgressBar&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ProgressBar::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString ProgressBar::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ProgressBar::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP ProgressBar::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_PROGBAR_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_PROGBAR_BARSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.barStyle), description);
			break;
		case DISPID_PROGBAR_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_PROGBAR_HALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hAlignment), description);
			break;
		case DISPID_PROGBAR_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_PROGBAR_ORIENTATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.orientation), description);
			break;
		case DISPID_PROGBAR_PROGRESSSTATE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.progressState), description);
			break;
		case DISPID_PROGBAR_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_PROGBAR_TEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		default:
			return IPerPropertyBrowsingImpl<ProgressBar>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ProgressBar::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_PROGBAR_BORDERSTYLE:
		case DISPID_PROGBAR_ORIENTATION:
			c = 2;
			break;
		case DISPID_PROGBAR_BARSTYLE:
		case DISPID_PROGBAR_HALIGNMENT:
		case DISPID_PROGBAR_PROGRESSSTATE:
			c = 3;
			break;
		case DISPID_PROGBAR_APPEARANCE:
		case DISPID_PROGBAR_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_PROGBAR_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<ProgressBar>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_PROGBAR_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		} else if(property == DISPID_PROGBAR_PROGRESSSTATE) {
			// the enum is 1-based
			++propertyValue;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_PROGBAR_APPEARANCE:
		case DISPID_PROGBAR_BARSTYLE:
		case DISPID_PROGBAR_BORDERSTYLE:
		case DISPID_PROGBAR_HALIGNMENT:
		case DISPID_PROGBAR_MOUSEPOINTER:
		case DISPID_PROGBAR_ORIENTATION:
		case DISPID_PROGBAR_PROGRESSSTATE:
		case DISPID_PROGBAR_RIGHTTOLEFT:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<ProgressBar>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_PROGBAR_TEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<ProgressBar>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ProgressBar::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_PROGBAR_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				case aDefault:
					description = GetResStringWithNumber(IDP_APPEARANCEDEFAULT, aDefault);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PROGBAR_BARSTYLE:
			switch(cookie) {
				case basSegments:
					description = GetResStringWithNumber(IDP_BARSTYLESEGMENTS, basSegments);
					break;
				case basSmooth:
					description = GetResStringWithNumber(IDP_BARSTYLESMOOTH, basSmooth);
					break;
				case basMarquee:
					description = GetResStringWithNumber(IDP_BARSTYLEMARQUEE, basMarquee);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PROGBAR_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PROGBAR_HALIGNMENT:
			switch(cookie) {
				case halLeft:
					description = GetResStringWithNumber(IDP_HALIGNMENTLEFT, halLeft);
					break;
				case halCenter:
					description = GetResStringWithNumber(IDP_HALIGNMENTCENTER, halCenter);
					break;
				case halRight:
					description = GetResStringWithNumber(IDP_HALIGNMENTRIGHT, halRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PROGBAR_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PROGBAR_ORIENTATION:
			switch(cookie) {
				case oHorizontal:
					description = GetResStringWithNumber(IDP_ORIENTATIONHORIZONTAL, oHorizontal);
					break;
				case oVertical:
					description = GetResStringWithNumber(IDP_ORIENTATIONVERTICAL, oVertical);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PROGBAR_PROGRESSSTATE:
			switch(cookie) {
				case psNormal:
					description = GetResStringWithNumber(IDP_PROGRESSSTATENORMAL, psNormal);
					break;
				case psPaused:
					description = GetResStringWithNumber(IDP_PROGRESSSTATEPAUSED, psPaused);
					break;
				case psFailed:
					description = GetResStringWithNumber(IDP_PROGRESSSTATEFAILED, psFailed);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PROGBAR_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP ProgressBar::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 5;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_StockColorPage;
		pPropertyPages->pElems[3] = CLSID_StockFontPage;
		pPropertyPages->pElems[4] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP ProgressBar::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ProgressBar>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ProgressBar::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP ProgressBar::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT ProgressBar::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT ProgressBar::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_PROGBAR_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT ProgressBar::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP ProgressBar::get_ActivateMarquee(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.activateMarquee);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_ActivateMarquee(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_ACTIVATEMARQUEE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.activateMarquee != b) {
		properties.activateMarquee = b;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(PBM_SETMARQUEE, properties.activateMarquee, properties.marqueeStepDuration);
			// NOTE: OnSetMarquee will call SetupTaskBar
			//SetupTaskBar();
		}
		FireOnChanged(DISPID_PROGBAR_ACTIVATEMARQUEE);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_APPEARANCE);
	if(newValue < 0 || newValue > aDefault) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(newValue == aDefault && !IsInDesignMode() && IsWindow()) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_PROGBAR_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 7;
	return S_OK;
}

STDMETHODIMP ProgressBar::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ProgressBar");
	return S_OK;
}

STDMETHODIMP ProgressBar::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ProgBar");
	return S_OK;
}

STDMETHODIMP ProgressBar::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF tmp;
		if(IsComctl32Version610OrNewer()) {
			tmp = SendMessage(PBM_GETBKCOLOR, 0, 0);
		} else {
			CWindowEx(*this).InternalSetRedraw(FALSE);
			tmp = SendMessage(PBM_SETBKCOLOR, 0, CLR_DEFAULT);
			SendMessage(PBM_SETBKCOLOR, 0, tmp);
			CWindowEx(*this).InternalSetRedraw(TRUE);
		}
		if(tmp == CLR_DEFAULT) {
			properties.backColor = CLR_NONE;
		} else if(tmp != OLECOLOR2COLORREF(properties.backColor)) {
			properties.backColor = tmp;
		}
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(PBM_SETBKCOLOR, 0, (properties.backColor == CLR_NONE ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.backColor)));
			FireViewChange();
		}
		FireOnChanged(DISPID_PROGBAR_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_BarColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF tmp;
		if(IsComctl32Version610OrNewer()) {
			tmp = SendMessage(PBM_GETBARCOLOR, 0, 0);
		} else {
			CWindowEx(*this).InternalSetRedraw(FALSE);
			tmp = SendMessage(PBM_SETBARCOLOR, 0, CLR_DEFAULT);
			SendMessage(PBM_SETBARCOLOR, 0, tmp);
			CWindowEx(*this).InternalSetRedraw(TRUE);
		}
		if(tmp == CLR_DEFAULT) {
			properties.barColor = CLR_NONE;
		} else if(tmp != OLECOLOR2COLORREF(properties.barColor)) {
			properties.barColor = tmp;
		}
	}

	*pValue = properties.barColor;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_BarColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_BARCOLOR);
	if(properties.barColor != newValue) {
		properties.barColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(PBM_SETBARCOLOR, 0, (properties.barColor == CLR_NONE ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.barColor)));
			FireViewChange();
		}
		FireOnChanged(DISPID_PROGBAR_BARCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_BarStyle(BarStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BarStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		if(style & PBS_SMOOTH) {
			properties.barStyle = basSmooth;
		} else if(RunTimeHelper::IsCommCtrl6() && (style & PBS_MARQUEE) == PBS_MARQUEE) {
			properties.barStyle = basMarquee;
		} else {
			properties.barStyle = basSegments;
		}
	}

	*pValue = properties.barStyle;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_BarStyle(BarStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_BARSTYLE);
	if(properties.barStyle != newValue) {
		properties.barStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(RunTimeHelper::IsCommCtrl6()) {
				properties.currentValue = SendMessage(PBM_GETPOS, 0, 0);
				switch(properties.barStyle) {
					case basSegments:
						ModifyStyle(PBS_SMOOTH | PBS_MARQUEE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
					case basSmooth:
						ModifyStyle(PBS_MARQUEE, PBS_SMOOTH, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
					case basMarquee:
						ModifyStyle(PBS_SMOOTH, PBS_MARQUEE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
				}
				SendMessage(PBM_SETPOS, properties.currentValue, 0);
				SetupTaskBar();
			} else {
				RecreateControlWindow();
			}
		}
		FireOnChanged(DISPID_PROGBAR_BARSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_PROGBAR_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ProgressBar::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP ProgressBar::get_CurrentValue(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.currentValue = SendMessage(PBM_GETPOS, 0, 0);
	}

	*pValue = properties.currentValue;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_CurrentValue(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_CURRENTVALUE);
	if(properties.currentValue != newValue) {
		properties.currentValue = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(PBM_SETPOS, properties.currentValue, 0);
			SetupTaskBar();
		}
		FireOnChanged(DISPID_PROGBAR_CURRENTVALUE);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_DetectDoubleClicks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_DetectDoubleClicks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_DETECTDOUBLECLICKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.detectDoubleClicks != b) {
		properties.detectDoubleClicks = b;
		SetDirty(TRUE);

		FireOnChanged(DISPID_PROGBAR_DETECTDOUBLECLICKS);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deMouseEvents) != (newValue & deMouseEvents)) {
			if(IsWindow()) {
				if(newValue & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_PROGBAR_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_DisplayProgressInTaskBar(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayProgressInTaskBar);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_DisplayProgressInTaskBar(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_DISPLAYPROGRESSINTASKBAR);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayProgressInTaskBar != b) {
		properties.displayProgressInTaskBar = b;
		SetDirty(TRUE);

		if(properties.pTaskBarList) {
			if(::IsWindow(properties.hApplicationWindow)) {
				ATLVERIFY(SUCCEEDED(properties.pTaskBarList->SetProgressState(properties.hApplicationWindow, TBPF_NOPROGRESS)));
			}
			properties.pTaskBarList->Release();
			properties.pTaskBarList = NULL;
		}
		if(properties.displayProgressInTaskBar) {
			if(!IsInDesignMode() && IsWindow()) {
				if(SUCCEEDED(CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&properties.pTaskBarList)))) {
					HRESULT hr = properties.pTaskBarList->HrInit();
					ATLASSERT(SUCCEEDED(hr));
					ATLASSUME(properties.pTaskBarList);
					if(FAILED(hr)) {
						properties.pTaskBarList->Release();
						properties.pTaskBarList = NULL;
					}
				}
			}
		}
		SetupTaskBar();

		FireOnChanged(DISPID_PROGBAR_DISPLAYPROGRESSINTASKBAR);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_DisplayText(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayText);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_DisplayText(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_DISPLAYTEXT);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayText != b) {
		properties.displayText = b;
		SetDirty(TRUE);
		FireViewChange();
		FireOnChanged(DISPID_PROGBAR_DISPLAYTEXT);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_PROGBAR_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_PROGBAR_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_PROGBAR_FONT);

	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_PROGBAR_FONT);
	return S_OK;
}

STDMETHODIMP ProgressBar::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_PROGBAR_FONT);

	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_PROGBAR_FONT);
	return S_OK;
}

STDMETHODIMP ProgressBar::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_PROGBAR_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_HAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hAlignment;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_HAlignment(HAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_HALIGNMENT);
	if(properties.hAlignment != newValue) {
		properties.hAlignment = newValue;
		SetDirty(TRUE);
		FireViewChange();
		FireOnChanged(DISPID_PROGBAR_HALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_hApplicationWindow(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.hApplicationWindow);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_hApplicationWindow(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_HAPPLICATIONWINDOW);
	HWND hApplicationWindow = static_cast<HWND>(LongToHandle(newValue));
	if(properties.hApplicationWindow != hApplicationWindow) {
		if(properties.pTaskBarList && ::IsWindow(properties.hApplicationWindow)) {
			ATLVERIFY(SUCCEEDED(properties.pTaskBarList->SetProgressState(properties.hApplicationWindow, TBPF_NOPROGRESS)));
		}
		properties.hApplicationWindow = hApplicationWindow;
		SetDirty(TRUE);

		SetupTaskBar();
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_PROGBAR_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP ProgressBar::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP ProgressBar::get_MarqueeStepDuration(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.marqueeStepDuration;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_MarqueeStepDuration(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_MARQUEESTEPDURATION);
	if(properties.marqueeStepDuration != newValue) {
		properties.marqueeStepDuration = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			SendMessage(PBM_SETMARQUEE, properties.activateMarquee, properties.marqueeStepDuration);
		}
		FireOnChanged(DISPID_PROGBAR_MARQUEESTEPDURATION);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Maximum(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.maximum = SendMessage(PBM_GETRANGE, FALSE, 0);
	}

	*pValue = properties.maximum;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_Maximum(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_MAXIMUM);
	if(properties.maximum != newValue) {
		properties.maximum = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			int minimum = SendMessage(PBM_GETRANGE, TRUE, 0);
			SendMessage(PBM_SETRANGE32, minimum, properties.maximum);
			SetupTaskBar();
		}
		FireOnChanged(DISPID_PROGBAR_MAXIMUM);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Minimum(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.minimum = SendMessage(PBM_GETRANGE, TRUE, 0);
	}

	*pValue = properties.minimum;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_Minimum(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_MINIMUM);
	if(properties.minimum != newValue) {
		properties.minimum = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			int maximum = SendMessage(PBM_GETRANGE, FALSE, 0);
			SendMessage(PBM_SETRANGE32, properties.minimum, maximum);
			SetupTaskBar();
		}
		FireOnChanged(DISPID_PROGBAR_MINIMUM);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_PROGBAR_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_PROGBAR_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ProgressBar::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_PROGBAR_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_PROGBAR_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ProgressBar::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_PROGBAR_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Orientation(OrientationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OrientationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.orientation = ((GetStyle() & PBS_VERTICAL) == PBS_VERTICAL ? oVertical : oHorizontal);
	}

	*pValue = properties.orientation;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_Orientation(OrientationConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_ORIENTATION);
	if(properties.orientation != newValue) {
		properties.orientation = newValue;
		SetDirty(TRUE);

		RECT windowRect = m_rcPos;
		SIZEL himetric = {m_sizeExtent.cx, m_sizeExtent.cy};
		SIZEL pixels = {0};
	  AtlHiMetricToPixel(&himetric, &pixels);
		windowRect.right = windowRect.left + pixels.cy;
		windowRect.bottom = windowRect.top + pixels.cx;
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite) {
			m_spInPlaceSite->OnPosRectChange(&windowRect);
		}

		if(IsWindow()) {
			if(RunTimeHelper::IsCommCtrl6()) {
				switch(properties.orientation) {
					case oHorizontal:
						ModifyStyle(PBS_VERTICAL, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
					case oVertical:
						ModifyStyle(0, PBS_VERTICAL, SWP_DRAWFRAME | SWP_FRAMECHANGED);
						break;
				}
				FireViewChange();
			} else {
				RecreateControlWindow();
			}
		}
		FireOnChanged(DISPID_PROGBAR_ORIENTATION);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ProgressBar::get_ProgressState(ProgressStateConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ProgressStateConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.progressState = static_cast<ProgressStateConstants>(SendMessage(PBM_GETSTATE, 0, 0));
	}

	*pValue = properties.progressState;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_ProgressState(ProgressStateConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_PROGRESSSTATE);
	if(properties.progressState != newValue) {
		properties.progressState = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			SendMessage(PBM_SETSTATE, properties.progressState, 0);
			FireViewChange();
			SetupTaskBar();
		}
		FireOnChanged(DISPID_PROGBAR_PROGRESSSTATE);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
		}
		FireOnChanged(DISPID_PROGBAR_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetExStyle();
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.rightToLeft & rtlLayout) {
				if(properties.rightToLeft & rtlText) {
					ModifyStyleEx(WS_EX_LTRREADING, WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT | WS_EX_RTLREADING);
				} else {
					ModifyStyleEx(WS_EX_RTLREADING, WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT | WS_EX_LTRREADING);
				}
			} else {
				if(properties.rightToLeft & rtlText) {
					ModifyStyleEx(WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT | WS_EX_LTRREADING, WS_EX_RTLREADING);
				} else {
					ModifyStyleEx(WS_EX_LAYOUTRTL | WS_EX_NOINHERITLAYOUT | WS_EX_RTLREADING, WS_EX_LTRREADING);
				}
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_PROGBAR_RIGHTTOLEFT);
		FireOnChanged(DISPID_PROGBAR_RIGHTTOLEFTLAYOUT);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_RightToLeftLayout(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(0);
	get_RightToLeft(&rightToLeft);
	*pValue = BOOL2VARIANTBOOL(rightToLeft & rtlLayout);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_RightToLeftLayout(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_RIGHTTOLEFTLAYOUT);
	BOOL b = VARIANTBOOL2BOOL(newValue);
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(0);
	get_RightToLeft(&rightToLeft);
	if(((rightToLeft & rtlLayout) == rtlLayout) != b) {
		if(b) {
			rightToLeft = static_cast<RightToLeftConstants>(rightToLeft | rtlLayout);
		} else {
			rightToLeft = static_cast<RightToLeftConstants>(rightToLeft & ~rtlLayout);
		}
		put_RightToLeft(rightToLeft);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_SmoothReverse(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.smoothReverse = ((GetStyle() & PBS_SMOOTHREVERSE) == PBS_SMOOTHREVERSE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.smoothReverse);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_SmoothReverse(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_SMOOTHREVERSE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.smoothReverse != b) {
		properties.smoothReverse = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			int currentPosition = SendMessage(PBM_GETPOS, 0, 0);
			if(properties.smoothReverse) {
				ModifyStyle(0, PBS_SMOOTHREVERSE);
			} else {
				ModifyStyle(PBS_SMOOTHREVERSE, 0);
			}
			// Windows has reset the position to 0
			SendMessage(PBM_SETPOS, currentPosition, 0);
		}
		FireOnChanged(DISPID_PROGBAR_SMOOTHREVERSE);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_StepWidth(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.stepWidth = SendMessage(PBM_GETSTEP, 0, 0);
	}

	*pValue = properties.stepWidth;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_StepWidth(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_STEPWIDTH);
	if(properties.stepWidth != newValue) {
		properties.stepWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(PBM_SETSTEP, properties.stepWidth, 0);
		}
		FireOnChanged(DISPID_PROGBAR_STEPWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_PROGBAR_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP ProgressBar::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_TEXT);
	if(properties.text != newValue) {
		properties.text.AssignBSTR(newValue);
		SetDirty(TRUE);
		FireViewChange();
		FireOnChanged(DISPID_PROGBAR_TEXT);
		SendOnDataChange();
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_TextShadowColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.textShadowColor;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_TextShadowColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_TEXTSHADOWCOLOR);
	if(properties.textShadowColor != newValue) {
		properties.textShadowColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_PROGBAR_TEXTSHADOWCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_TextShadowOffsetX(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.textShadowOffsetX;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_TextShadowOffsetX(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_TEXTSHADOWOFFSETX);
	if(properties.textShadowOffsetX != newValue) {
		properties.textShadowOffsetX = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_PROGBAR_TEXTSHADOWOFFSETX);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_TextShadowOffsetY(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.textShadowOffsetY;
	return S_OK;
}

STDMETHODIMP ProgressBar::put_TextShadowOffsetY(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_TEXTSHADOWOFFSETY);
	if(properties.textShadowOffsetY != newValue) {
		properties.textShadowOffsetY = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_PROGBAR_TEXTSHADOWOFFSETY);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP ProgressBar::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PROGBAR_USESYSTEMFONT);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_PROGBAR_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP ProgressBar::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ProgressBar::ChangeCurrentValue(LONG delta/* = 0*/, LONG* pPreviousValue/* = NULL*/)
{
	ATLASSERT_POINTER(pPreviousValue, LONG);
	if(!pPreviousValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		if(delta == 0) {
			*pPreviousValue = SendMessage(PBM_STEPIT, 0, 0);
		} else {
			*pPreviousValue = SendMessage(PBM_DELTAPOS, delta, 0);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ProgressBar::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP ProgressBar::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		UpdateWindow();
	}
	return S_OK;
}

STDMETHODIMP ProgressBar::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}


LRESULT ProgressBar::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);
	return 0;
}

LRESULT ProgressBar::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT ProgressBar::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT ProgressBar::OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL inDesignMode = IsInDesignMode();
	if(!inDesignMode) {
		TCHAR pBuffer[200];
		ZeroMemory(pBuffer, 200 * sizeof(TCHAR));
		GetModuleFileName(NULL, pBuffer, 200);
		PathStripPath(pBuffer);
		if(lstrcmpi(pBuffer, TEXT("vb6.exe")) == 0) {
			HWND hWnd = GetParent();
			while(hWnd) {
				if(GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("DesignerWindow")) == 0) {
					hWnd = ::GetParent(hWnd);
					if(hWnd && GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("MDIClient")) == 0) {
						hWnd = ::GetParent(hWnd);
						if(hWnd && GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("wndclass_desked_gsk")) == 0) {
							inDesignMode = TRUE;
						}
					}
					break;
				}
				hWnd = ::GetParent(hWnd);
			}
		}
	}
	if(!inDesignMode) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		if(WindowFromPoint(mousePosition) == *this) {
			ScreenToClient(&mousePosition);
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			if(HIWORD(lParam) == WM_LBUTTONDOWN) {
				button = 1/*MouseButtonConstants.vbLeftButton*/;
				Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
			} else if(HIWORD(lParam) == WM_MBUTTONDOWN) {
				button = 4/*MouseButtonConstants.vbMiddleButton*/;
				Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
			} else if(HIWORD(lParam) == WM_RBUTTONDOWN) {
				button = 2/*MouseButtonConstants.vbRightButton*/;
				Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
			} else if(HIWORD(lParam) == WM_XBUTTONDOWN) {
				Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
			}
			return MA_NOACTIVATEANDEAT;
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ProgressBar::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);

	return 0;
}

LRESULT ProgressBar::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ProgressBar::OnNCHitTest(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	return HTCLIENT;
}

LRESULT ProgressBar::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	HRGN hRegion = NULL;
	BOOL drawText = (properties.displayText && APIWrapper::IsSupported_DrawShadowText() && properties.text && properties.text.Length() > 0);

	if(drawText && message == WM_PAINT) {
		hRegion = CreateRectRgn(0, 0, 0, 0);
		if(hRegion) {
			if(GetUpdateRgn(hRegion, FALSE) == ERROR) {
				DeleteObject(hRegion);
				hRegion = NULL;
			}
		}
	}
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	// TODO: reduce flicker
	if(drawText) {
		WTL::CString templateText = properties.text;
		int value = SendMessage(PBM_GETPOS, 0, 0);
		int range = SendMessage(PBM_GETRANGE, FALSE, 0) - SendMessage(PBM_GETRANGE, TRUE, 0);
		int percent = 0;
		if(range != 0) {
			percent = static_cast<int>(static_cast<__int64>(value) * static_cast<__int64>(100) / static_cast<__int64>(range));
		}

		WTL::CString text;
		for(int i = 0; i < templateText.GetLength(); i++) {
			if(templateText[i] == TEXT('%')) {
				if(templateText[i + 1] == TEXT('%')) {
					text += templateText[i];
					i++;
				} else if(templateText[i + 1] == TEXT('1')) {
					WTL::CString s;
					s.Format(TEXT("%i"), percent);
					text += s;
					i++;
				} else if(templateText[i + 1] == TEXT('2')) {
					WTL::CString s;
					s.Format(TEXT("%i"), value);
					text += s;
					i++;
				} else {
					text += templateText[i];
				}
			} else {
				text += templateText[i];
			}
		}

		CT2CW converter(text);
		LPCWSTR pText = converter;
		int textLength = 0;
		if(pText) {
			textLength = lstrlenW(pText);
		}
		if(textLength > 0) {
			HDC hDC = NULL;
			if(message == WM_PRINTCLIENT) {
				hDC = reinterpret_cast<HDC>(wParam);
			} else {
				hDC = GetDC();
			}
			if(hDC) {
				if(hRegion) {
					SelectClipRgn(hDC, hRegion);
				}
				RECT clientRectangle = {0};
				GetClientRect(&clientRectangle);
				clientRectangle.left += 4;
				clientRectangle.right -= 4;

				HGDIOBJ hOldFont = SelectObject(hDC, reinterpret_cast<HGDIOBJ>(SendMessage(WM_GETFONT, 0, 0)));
				WTL::CRect textRectangle = clientRectangle;
				DWORD drawTextFlags = DT_SINGLELINE | DT_EDITCONTROL | DT_NOPREFIX;
				if(GetExStyle() & WS_EX_RTLREADING) {
					drawTextFlags |= DT_RTLREADING;
				}
				DrawTextW(hDC, pText, textLength, &textRectangle, drawTextFlags | DT_CALCRECT);
				textRectangle.MoveToY((clientRectangle.top + clientRectangle.bottom - (textRectangle.bottom - textRectangle.top)) >> 1);
				textRectangle.left = clientRectangle.left;
				textRectangle.right = clientRectangle.right;

				switch(properties.hAlignment) {
					case halLeft:
						drawTextFlags |= DT_LEFT;
						break;
					case halCenter:
						drawTextFlags |= DT_CENTER;
						break;
					case halRight:
						drawTextFlags |= DT_RIGHT;
						break;
				}
				APIWrapper::DrawShadowText(hDC, pText, textLength, &textRectangle, drawTextFlags, OLECOLOR2COLORREF(properties.foreColor), OLECOLOR2COLORREF(properties.textShadowColor), properties.textShadowOffsetX, properties.textShadowOffsetY, NULL);

				SelectObject(hDC, hOldFont);
				if(hRegion) {
					SelectClipRgn(hDC, NULL);
				}
				if(message != WM_PRINTCLIENT) {
					ReleaseDC(hDC);
				}
			}
		}
	}
	if(hRegion) {
		DeleteObject(hRegion);
	}
	return lr;
}

LRESULT ProgressBar::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	WTL::CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT ProgressBar::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_PROGBAR_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_PROGBAR_FONT);
	}

	return lr;
}

LRESULT ProgressBar::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT ProgressBar::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETNONCLIENTMETRICS) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ProgressBar::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ProgressBar::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	WTL::CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ProgressBar::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ProgressBar::OnSetMarquee(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	properties.activateMarquee = wParam;
	properties.marqueeStepDuration = lParam;
	SetupTaskBar();
	return lr;
}

LRESULT ProgressBar::OnSetStep(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	long oldValue = properties.stepWidth;
	properties.stepWidth = wParam;
	return oldValue;
}


void ProgressBar::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
			properties.font.owningFontResource = FALSE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT ProgressBar::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		/*if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
		}*/

		return Fire_ContextMenu(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	BOOL fireMouseDown = FALSE;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		fireMouseDown = TRUE;
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
	}
	mouseStatus.StoreClickCandidate(button);
	SetCapture();

	if(mouseStatus.IsDoubleClickCandidate(button)) {
		/* Watch for double-clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the double-click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_DblClick(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RDblClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MDblClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XDblClick(button, shift, x, y);
					}
					break;
			}
			mouseStatus.RemoveAllDoubleClickCandidates();
		}

		mouseStatus.RemoveClickCandidate(button);
		if(GetCapture() == *this) {
			ReleaseCapture();
		}
		fireMouseDown = FALSE;
	} else {
		mouseStatus.RemoveAllDoubleClickCandidates();
	}

	HRESULT hr = S_OK;
	if(fireMouseDown) {
		hr = Fire_MouseDown(button, shift, x, y);
	}
	return hr;
}

inline HRESULT ProgressBar::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseEnter(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT ProgressBar::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.LeaveControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT ProgressBar::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y);
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(GetCapture() == *this) {
			ReleaseCapture();
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
			if(properties.detectDoubleClicks) {
				mouseStatus.StoreDoubleClickCandidate(button);
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT ProgressBar::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

		if(pData) {
			/* Actually we wouldn't need the next line, because the data object passed to this method should
				 always be the same as the data object that was passed to Raise_OLEDragEnter. */
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT ProgressBar::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			return Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}
	return S_OK;
}

inline HRESULT ProgressBar::Raise_OLEDragLeave(void)
{
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);

		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT ProgressBar::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			return Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}
	return S_OK;
}

inline HRESULT ProgressBar::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}


void ProgressBar::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

inline HRESULT ProgressBar::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT ProgressBar::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y);
	//}
	//return S_OK;
}


DWORD ProgressBar::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD ProgressBar::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	switch(properties.barStyle) {
		case basSmooth:
			style |= PBS_SMOOTH;
			break;
		case basMarquee:
			if(RunTimeHelper::IsCommCtrl6()) {
				style |= PBS_MARQUEE;
			}
			break;
	}
	switch(properties.orientation) {
		case oVertical:
			style |= PBS_VERTICAL;
			break;
	}
	if(properties.smoothReverse) {
		if(IsComctl32Version610OrNewer()) {
			style |= PBS_SMOOTHREVERSE;
		}
	}
	return style;
}

void ProgressBar::SendConfigurationMessages(void)
{
	/* HACK: Msctls_progress32 will remove WS_EX_CLIENTEDGE on creation, but we can set it after the window
	         was created. */
	switch(properties.appearance) {
		case a2D:
			ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
		case a3D:
			ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
		case a3DLight:
			ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
			break;
	}
	if(properties.pTaskBarList) {
		properties.pTaskBarList->Release();
		properties.pTaskBarList = NULL;
	}
	if(properties.displayProgressInTaskBar) {
		if(!IsInDesignMode() && IsWindow()) {
			if(SUCCEEDED(CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&properties.pTaskBarList)))) {
				HRESULT hr = properties.pTaskBarList->HrInit();
				ATLASSERT(SUCCEEDED(hr));
				ATLASSUME(properties.pTaskBarList);
				if(FAILED(hr)) {
					properties.pTaskBarList->Release();
					properties.pTaskBarList = NULL;
				}
			}
		}
	}

	SendMessage(PBM_SETBKCOLOR, 0, (properties.backColor == CLR_NONE ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.backColor)));
	SendMessage(PBM_SETBARCOLOR, 0, (properties.barColor == CLR_NONE ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.barColor)));
	SendMessage(PBM_SETRANGE32, properties.minimum, properties.maximum);
	SendMessage(PBM_SETSTEP, properties.stepWidth, 0);
	SendMessage(PBM_SETPOS, properties.currentValue, 0);
	if(RunTimeHelper::IsCommCtrl6()) {
		SendMessage(PBM_SETMARQUEE, properties.activateMarquee, properties.marqueeStepDuration);
	}
	if(IsComctl32Version610OrNewer()) {
		SendMessage(PBM_SETSTATE, properties.progressState, 0);
	}

	ApplyFont();
	SetupTaskBar();
}

HCURSOR ProgressBar::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

void ProgressBar::SetupTaskBar(void)
{
	if(properties.displayProgressInTaskBar && properties.pTaskBarList && ::IsWindow(properties.hApplicationWindow)) {
		BarStyleConstants barStyle = basSegments;
		ATLVERIFY(SUCCEEDED(get_BarStyle(&barStyle)));
		if(barStyle == basMarquee) {
			VARIANT_BOOL activateMarquee = VARIANT_FALSE;
			ATLVERIFY(SUCCEEDED(get_ActivateMarquee(&activateMarquee)));
			if(VARIANTBOOL2BOOL(activateMarquee)) {
				ATLVERIFY(SUCCEEDED(properties.pTaskBarList->SetProgressState(properties.hApplicationWindow, TBPF_INDETERMINATE)));
			} else {
				ATLVERIFY(SUCCEEDED(properties.pTaskBarList->SetProgressState(properties.hApplicationWindow, TBPF_NOPROGRESS)));
			}
		} else {
			ProgressStateConstants progressState = psNormal;
			ATLVERIFY(SUCCEEDED(get_ProgressState(&progressState)));
			switch(progressState) {
				case psNormal:
					ATLVERIFY(SUCCEEDED(properties.pTaskBarList->SetProgressState(properties.hApplicationWindow, TBPF_NORMAL)));
					break;
				case psFailed:
					ATLVERIFY(SUCCEEDED(properties.pTaskBarList->SetProgressState(properties.hApplicationWindow, TBPF_ERROR)));
					break;
				case psPaused:
					ATLVERIFY(SUCCEEDED(properties.pTaskBarList->SetProgressState(properties.hApplicationWindow, TBPF_PAUSED)));
					break;
			}

			LONG minimum = 0;
			LONG maximum = 0;
			LONG currentValue = 0;
			ATLVERIFY(SUCCEEDED(get_Minimum(&minimum)));
			ATLVERIFY(SUCCEEDED(get_Maximum(&maximum)));
			ATLVERIFY(SUCCEEDED(get_CurrentValue(&currentValue)));
			ATLVERIFY(SUCCEEDED(properties.pTaskBarList->SetProgressValue(properties.hApplicationWindow, currentValue - minimum, maximum - minimum)));
		}
	}
}

BOOL ProgressBar::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}


BOOL ProgressBar::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}