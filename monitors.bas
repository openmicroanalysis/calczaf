Attribute VB_Name = "CodeMonitors"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private Type TypeRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TypeMONITORINFO
    cbSize As Long
    rcMonitor As TypeRECT
    rcWork As TypeRECT
    dwFlags As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As TypeRECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As TypeRECT, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwdata As Any) As Long
Private Declare Function MonitorFromRect Lib "user32" (ByRef lprc As TypeRECT, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As TypeMONITORINFO) As Long
Private Declare Function UnionRect Lib "user32" (lprcDst As TypeRECT, lprcSrc1 As TypeRECT, lprcSrc2 As TypeRECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Const MONITOR_DEFAULTTONEAREST = &H2

Dim rcMonitors() As TypeRECT    ' coordinate array for all monitors
Dim rcVS         As TypeRECT    ' coordinates for Virtual Screen

Private Function MonitorsEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As TypeRECT, dwdata As Long) As Long
    
ierror = False
On Error GoTo MonitorsEnumProcError
    
    ReDim Preserve rcMonitors(dwdata)
    
    rcMonitors(dwdata) = lprcMonitor
    UnionRect rcVS, rcVS, lprcMonitor   ' merge all monitors together to get the virtual screen coordinates
    dwdata = dwdata + 1                 ' increase monitor count
    MonitorsEnumProc = 1                ' continue

Exit Function

' Errors
MonitorsEnumProcError:
MsgBox Error$, vbOKOnly + vbCritical, "MonitorsEnumProc"
ierror = True
Exit Function

End Function

Sub MonitorsSavePosition(hWnd As Long, string1 As String, string2 As String)
' Save the form position for the form handle passed
    
ierror = False
On Error GoTo MonitorsSavePositionError
    
    Dim rc As TypeRECT
    GetWindowRect hWnd&, rc ' save position in pixel units
    SaveSetting string1$, string2$, "Left", rc.Left
    SaveSetting string1$, string2$, "Top", rc.Top

Exit Sub

' Errors
MonitorsSavePositionError:
MsgBox Error$, vbOKOnly + vbCritical, "MonitorsSavePosition"
ierror = True
Exit Sub

End Sub

Sub MonitorsLoadPosition(hWnd As Long, string1 As String, string2 As String)
' Load the form position for the form handle passed
    
ierror = False
On Error GoTo MonitorsLoadPositionError
    
    Dim rc As TypeRECT, Left As Long, Top As Long, hMonitor As Long, mi As TypeMONITORINFO
    
    GetWindowRect hWnd, rc ' obtain the window rectangle
    
    ' Move the window rectangle to position saved previously
    Left = GetSetting(string1$, string2$, "Left", rc.Left)
    Top = GetSetting(string1$, string2$, "Top", rc.Left)
    OffsetRect rc, Left - rc.Left, Top - rc.Top
    
    ' Find the monitor closest to window rectangle
    hMonitor = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
    
    ' Get info about monitor coordinates and working area
    mi.cbSize = Len(mi)
    GetMonitorInfo hMonitor, mi
    
    ' Adjust the window rectangle so it fits inside the work area of the monitor
    If rc.Left < mi.rcWork.Left Then OffsetRect rc, mi.rcWork.Left - rc.Left, 0
    If rc.Right > mi.rcWork.Right Then OffsetRect rc, mi.rcWork.Right - rc.Right, 0
    If rc.Top < mi.rcWork.Top Then OffsetRect rc, 0, mi.rcWork.Top - rc.Top
    If rc.Bottom > mi.rcWork.Bottom Then OffsetRect rc, 0, mi.rcWork.Bottom - rc.Bottom
    
    ' Move the window to new calculated position
    MoveWindow hWnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 0

Exit Sub

' Errors
MonitorsLoadPositionError:
MsgBox Error$, vbOKOnly + vbCritical, "MonitorsLoadPosition"
ierror = True
Exit Sub

End Sub

Sub MonitorsEnumMonitors(f As Form)
' Code to display test form with monitor info

ierror = False
On Error GoTo MonitorsEnumMonitorsError

    Dim n As Long
    
    ' Get the monitor info
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorsEnumProc, n&
    
    With f
        .Move .Left, .Top, (rcVS.Right - rcVS.Left) * 2 + .Width - .ScaleWidth, (rcVS.Bottom - rcVS.Top) * 2 + .Height - .ScaleHeight
    End With
    
    f.Scale (rcVS.Left, rcVS.Top)-(rcVS.Right, rcVS.Bottom)
    f.Caption = n & " Monitor" & IIf(n > 1, "s", vbNullString)
    f.LblMonitors(0).Appearance = 0 'Flat
    f.LblMonitors(0).BorderStyle = 1 'FixedSingle
    
    For n = 0 To n - 1
        If n Then
            Load f.LblMonitors(n)
            f.LblMonitors(n).Visible = True
        End If
        
        With rcMonitors(n)
            f.LblMonitors(n).Move .Left, .Top, .Right - .Left, .Bottom - .Top
            f.LblMonitors(n).Caption = "Monitor " & n + 1 & vbLf & _
                .Right - .Left & " x " & .Bottom - .Top & vbLf & _
                "(" & .Left & ", " & .Top & ")-(" & .Right & ", " & .Bottom & ")"
        End With
    Next

Exit Sub

' Errors
MonitorsEnumMonitorsError:
MsgBox Error$, vbOKOnly + vbCritical, "MonitorsEnumMonitors"
ierror = True
Exit Sub

End Sub

Sub MonitorsGetVirtualExtents(nMonitors As Long, tWidth() As Long, tHeight() As Long, vWidth As Long, vHeight As Long)
' Code to return full virtual extent of display area (for two monitors)

ierror = False
On Error GoTo MonitorsGetVirtualExtentsError

    Dim n As Long, nn As Long
    
    ' Get the monitor info
    EnumDisplayMonitors 0&, ByVal 0&, AddressOf MonitorsEnumProc, nn&
    
    ' Return virtual width and height
    vWidth& = rcVS.Right - rcVS.Left
    vHeight& = rcVS.Bottom - rcVS.Top
    
    ' Dimension return arrays of each monitor
    ReDim tWidth(1 To nn&) As Long
    ReDim tHeight(1 To nn&) As Long
        
    For n& = 0 To nn& - 1
        With rcMonitors(n&)
            tWidth&(n& + 1) = .Right - .Left
            tHeight&(n& + 1) = .Bottom - .Top
        End With
    Next

nMonitors& = nn&
Exit Sub

' Errors
MonitorsGetVirtualExtentsError:
MsgBox Error$, vbOKOnly + vbCritical, "MonitorsGetVirtualExtents"
ierror = True
Exit Sub

End Sub





