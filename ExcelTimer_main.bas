Attribute VB_Name = "ExcelTimer_main"
#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Dim Count As Long 'Timerループの管理
Dim Flashing As Byte
Dim StopSignal As Byte '0ならTimer停止、1ならリセット
Dim Button As Byte '1になるとTimerループが停止する
Dim ResetSignal As Byte
Dim CountProhibition As Byte '手動カウントの禁止
'更新日2022-11-22

Sub Timer()
    
    Range("A1").Select

    Count = 0
    Button = 0
    StopSignal = 0
    CountProhibition = 1
    
    For Count = 0 To 9999999 Step 1
    
        DoEvents
            
            If Button = 1 Then
                Range("J1") = ""
                StopSignal = 1
                CountProhibition = 0
                Exit Sub
                
            End If
            
            Range("J1") = "カウント中・・・"
            
            If Range("D2") < 60 Then
            
                Sleep 1000
                
                Range("D2") = Range("D2") + 1

                
            ElseIf Range("D2") = 60 Then
                
                Range("B2") = Range("B2") + 1
                Range("D2") = Range("D2") - Range("D2")
                
            Else
            
                MsgBox "範囲エラー", vbCritical, Title:="ERROR"
                
            End If
            
            If Count = 9999999 Then
            
                Count = Count - Count
            Else
            
    Next Count
    
End Sub

Sub eButton()

    Range("A1").Select

    If StopSignal = 0 Then
        
        Button = 1
        
    End If

    If StopSignal = 1 Then

        Call Reset
        
    End If
    
End Sub

Sub Reset()

    Range("A1").Select

    If Not Range("B2") = 0 Then
    
        Range("B2") = 0
        
    End If
    
    
    If Not Range("D2") = 0 Then
    
        Range("D2") = 0
    
    End If
    
    StopSignal = 0
    Button = 1

End Sub
Sub Plus_s()

    Range("A1").Select
    
    If CountProhibition = 0 Then
    
        If Range("D2") < 59 Then
        
            Range("D2") = Range("D2") + 1
            
        ElseIf Range("D2") = 59 Then
         
            Range("B2") = Range("B2") + 1
            Range("D2") = Range("D2") - Range("D2")
            
        Else
        
           MsgBox "範囲エラー", vbCritical, Title:="ERROR"
           
        End If
    
        StopSignal = 1
    
    Else
        
    End If
    
End Sub

Sub Minus_s()

    Range("A1").Select

    If CountProhibition = 0 Then

        If Range("D2") = 0 And Range("B2") = 0 Then
        
        ElseIf Range("D2") = 0 And Range("B2") >= 1 Then
            
            Range("B2") = Range("B2") - 1
            Range("D2") = 59
            
        ElseIf 1 <= Range("D2") And Range("D2") < 60 Then
        
            Range("D2") = Range("D2") - 1
        
        Else
            MsgBox "範囲エラー", vbCritical, Title:="ERROR"
        
        End If
    
        StopSignal = 1
    
    Else
        
    End If
    
End Sub
