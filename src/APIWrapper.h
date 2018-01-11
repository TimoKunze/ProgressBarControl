//////////////////////////////////////////////////////////////////////
/// \class APIWrapper
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Provides wrappers for API functions not available on all supported systems</em>
///
/// This class wraps calls to parts of the Win32 API that may be missing on the executing system.
//////////////////////////////////////////////////////////////////////


#pragma once


#ifndef DOXYGEN_SHOULD_SKIP_THIS
	typedef int WINAPI DrawShadowTextFn(__in HDC hdc, __in_ecount(cch) LPCWSTR pszText, __in UINT cch, __in RECT* prc, __in DWORD dwFlags, __in COLORREF crText, __in COLORREF crShadow, __in int ixOffset, __in int iyOffset);
#endif

class APIWrapper
{
private:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	APIWrapper(void);

public:
	/// \brief <em>Checks whether the executing system supports the \c DrawShadowText function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa DrawShadowText,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775639.aspx">DrawShadowText</a>
	static BOOL IsSupported_DrawShadowText(void);

	/// \brief <em>Calls the \c DrawShadowText function if it is available on the executing system</em>
	///
	/// \param[in] hDC The \c hdc parameter of the wrapped function.
	/// \param[in] pText The \c pszText parameter of the wrapped function.
	/// \param[in] textLength The \c cch parameter of the wrapped function.
	/// \param[in] pTextRectangle The \c prc parameter of the wrapped function.
	/// \param[in] flags The \c dwFlags parameter of the wrapped function.
	/// \param[in] textColor The \c crText parameter of the wrapped function.
	/// \param[in] shadowColor The \c crShadow parameter of the wrapped function.
	/// \param[in] xOffset The \c ixOffset parameter of the wrapped function.
	/// \param[in] yOffset The \c iyOffset parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_DrawShadowText,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775639.aspx">DrawShadowText</a>
	static HRESULT DrawShadowText(__in HDC hDC, __in_ecount(textLength) LPCWSTR pText, __in UINT textLength, __in RECT* pTextRectangle, __in DWORD flags, __in COLORREF textColor, __in COLORREF shadowColor, __in int xOffset, __in int yOffset, __deref_out_opt BOOL* pReturnValue);

protected:
	/// \brief <em>Stores whether support for the \c DrawShadowText API function has been checked</em>
	static BOOL checkedSupport_DrawShadowText;
	/// \brief <em>Caches the pointer to the \c DrawShadowText API function</em>
	static DrawShadowTextFn* pfnDrawShadowText;
};