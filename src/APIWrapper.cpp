// APIWrapper.cpp: A wrapper class for API functions not available on all supported systems

#include "stdafx.h"
#include "APIWrapper.h"


BOOL APIWrapper::checkedSupport_DrawShadowText = FALSE;
DrawShadowTextFn* APIWrapper::pfnDrawShadowText = NULL;


APIWrapper::APIWrapper(void)
{
}


BOOL APIWrapper::IsSupported_DrawShadowText(void)
{
	if(!checkedSupport_DrawShadowText) {
		HMODULE hComctl32 = GetModuleHandle(TEXT("comctl32.dll"));
		if(hComctl32) {
			pfnDrawShadowText = reinterpret_cast<DrawShadowTextFn*>(GetProcAddress(hComctl32, "DrawShadowText"));
		}
		checkedSupport_DrawShadowText = TRUE;
	}

	return (pfnDrawShadowText != NULL);
}

HRESULT APIWrapper::DrawShadowText(HDC hDC, LPCWSTR pText, UINT textLength, RECT* pTextRectangle, DWORD flags, COLORREF textColor, COLORREF shadowColor, int xOffset, int yOffset, BOOL* pReturnValue)
{
	BOOL dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_DrawShadowText()) {
		*pReturnValue = pfnDrawShadowText(hDC, pText, textLength, pTextRectangle, flags, textColor, shadowColor, xOffset, yOffset);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}