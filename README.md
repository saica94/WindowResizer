# Window Resizer
* [JP]ウィンドウのサイズを最大4K(3840x2160)まで拡大したり、最小10x10まで縮小できます。  
* [EN]You can enlarge the window size up to 4K (3840x2160) or reduce it to a minimum of 10x10.  

# Features
* [JP]使用しているモニタの解像度を超えてウィンドウサイズを拡大したい場合などに使えます。  
* [EN]This can be used when you want to enlarge the window size beyond the resolution of the monitor you are using.  

# Requirement
* Powershell 7
* .NET Core 6 ↑

# Installation
## PowerShell 7
* [winget]
``` PowerShell
 winget install --id Microsoft.Powershell --source winget
 ```
 
## .NET Core
* [winget]
``` PowerShell
winget install --id Microsoft.DotNet.SDK.8
```
### If unable to execute
Please check the current execution policy of PowerShell, and if it is set to Restricted, AllSigned, or RemoteSigned, change it to Unrestricted. When you are not using this program, remember to revert to the original execution policy.  

* Check the current execution policy:  
``` PowerShell
Get-ExecutionPolicy
```
* Change the execution policy:
``` PowerShell
Set-ExecutionPolicy Unrestricted -Focus CurrentUser -Force
```
* Revert to the original policy:
``` PowerShell
Set-ExecutionPolicy <Policy Level> -Focus CurrentUser -Force
```
※Policy Level  
| Policy Level | Local | Remote|
|--------------|-------|-------|
| Restricted | × | × |
| AllSigned | 〇(Signed only) | 〇(Signed only) |
| RemoteSigned | 〇 | 〇(Signed only)|
| Unrestricted | 〇 | 〇 |
# Usage
* [JP]ドロップダウンリストから変更したいウィンドウのタイトルを選択した後、WidthとHeightのスライダーを操作してウィンドウサイズを変更するか、入力ボックスに直接数値を入力してウィンドウサイズを変更してください。  
* [EN]After selecting the window title you want to change from the dropdown list, you can change the window size by manipulating the Width and Height sliders or by directly entering values into the input boxes.  

# Author
* Saica

# License
"Window Resizer" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).
