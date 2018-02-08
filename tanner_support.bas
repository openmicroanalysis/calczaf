Attribute VB_Name = "Tanner_SupportCode"
Option Explicit

Public Declare Function CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptrDst As Long, ByVal ptrSrc As Long, ByVal numOfBytes As Long) As Long

' Higher-performance timing functions are also handled by this class.  Note that you *must* initialize the timer engine
' before requesting any time values, or crashes will occurs because the frequency timer is 0.
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private m_TimerFrequency As Currency

Public Sub EnableHighResolutionTimers()
    QueryPerformanceFrequency m_TimerFrequency
    If m_TimerFrequency = 0 Then m_TimerFrequency = 1
End Sub

Public Function GetTimerDifference(ByRef startTime As Currency, ByRef stopTime As Currency) As Double
    GetTimerDifference = (stopTime - startTime) / m_TimerFrequency
End Function

Public Function GetTimerDifferenceNow(ByRef startTime As Currency) As Double
    Dim tmpTime As Currency
    QueryPerformanceCounter tmpTime
    GetTimerDifferenceNow = (tmpTime - startTime) / m_TimerFrequency
End Function

Public Sub GetHighResTime(ByRef dstTime As Currency)
    QueryPerformanceCounter dstTime
End Sub

Public Sub PrintTimeTakenInMs(ByRef startTime As Currency)
    Debug.Print "Time taken: " & Format$(Tanner_SupportCode.GetTimerDifferenceNow(startTime) * 1000, "####0.0") & " ms"
End Sub
