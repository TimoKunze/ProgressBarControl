<p align=center>
  <a href="https://github.com/TimoKunze/ProgressBarControl/releases/tag/2.6.3">
    <img alt="Release 2.6.3 Release" src="https://img.shields.io/badge/release-2.6.3-0688CB.svg">
  </a>
  <a href="https://github.com/TimoKunze/ProgressBarControl/releases">
    <img alt="Download ProgressBarControl" src="https://img.shields.io/badge/download-latest-0688CB.svg">
  </a>
  <a href="https://github.com/TimoKunze/ProgressBarControl/blob/master/LICENSE">
    <img alt="License" src="https://img.shields.io/badge/license-MIT-0688CB.svg">
  </a>
  <a href="https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ProgressBarControl&no_shipping=1&tax=0&currency_code=EUR">
    <img alt="Donate" src="https://img.shields.io/badge/%24-donate-E44E4A.svg">
  </a>
</p>

# ProgressBarControl
An ActiveX control for Visual Basic 6 that can be used to display a progress.

I've developed this ActiveX control in 2003 and did update it on a regular basis until 2016. Currently I have little interest to maintain this project any longer, but I think the code might be of some use to others.

# Before you make changes
If you make changes to the code and deploy the binary, keep in mind that ActiveX controls are COM components and therefore should stay binary compatible as long as you don't change the COM object's, i.e. the ActiveX control's public class name and GUIDs. Otherwise people using these components are likely to end up in the famous COM hell.

# Dependencies
You will need the [Microsoft Windows 10 SDK](https://developer.microsoft.com/en-us/windows/downloads/windows-10-sdk), ATL, and [WTL 10](https://sourceforge.net/projects/wtl/).

## Required Change to ATL
Some version of ATL have a bug in ```AtlIPersistPropertyBag_Load``` which causes crashes. In the file atlcom.h search for ```AtlIPersistPropertyBag_Load```. Inside the implementation of this function search for ```HRESULT hr = pPropBag->Read(pMap[i].szDesc, &var, pErrorLog);```. Make sure that there is this code before this line: ```var.pdispVal = NULL;```.