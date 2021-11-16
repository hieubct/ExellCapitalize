# ExellCapitalize
Capitalize all letters in a range with VBA code
Besides the formula method, you can run VBA code to capitalize all letters in a range.

#### 1. Press ALT + F11 keys simultaneously to open the Microsoft Visual Basic Application window.

#### 2. In the Microsoft Visual Basic Application window, click Insert > Module.

#### 3. Copy and paste below VBA code into the Module window.

VBA code: Capitalize all letters in a range
```VB	
Sub ToggleCase()
    Dim Rng As Range
    Dim WorkRng As Range
    On Error Resume Next
    xTitleId = "KutoolsforExcel"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type: = 8)
    For Each Rng In WorkRng
        Rng.Value = VBA.UCase(Rng.Value)
    Next
End Sub
```
#### 4. In the popping up dialog box, select the range with letters you want to capitalize, and then click the OK button. See screenshot:

Then all letters in selected range are all capitalized immediately.
![image](https://user-images.githubusercontent.com/43920196/142004839-4fcfb94d-29d2-4a5b-b3ef-64f9326e0668.png)
