Attribute VB_Name = "CodeMonitors"
' (c) Copyright 1995-2021 by John J. Donovan
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

Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hDC As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwdata As Any) As Long
Private Declare Function UnionRect Lib "user32" (lprcDst As TypeRECT, lprcSrc1 As TypeRECT, lprcSrc2 As TypeRECT) As Long

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





