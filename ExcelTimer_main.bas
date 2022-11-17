#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Dim Count As Long 'Timerループの管理
Dim Flashing As Byte
Dim StopSignal As Byte '0ならTimer停止、1ならリセット
Dim Button As Byte '1になるとTimerループが停止する
Dim CountProhibition As Byte '手動カウントの禁止

Sub Timer()

    Range("A1").Select

    Button = 0
    StopSignal = 0
    CountProhibition = 1
    
    For Count = 0 To 9999999
    
        DoEvents
            
            If Button = 1 Then
                Range("G1") = ""
                StopSignal = 1
                CountProhibition = 0
                Exit Sub
                
            End If
            
            Range("G1") = "カウント中・・・"
           
            If Range("D2") = 60 Then
                
                Range("B2") = Range("B2") + 1
                Range("D2") = Range("D2") - Range("D2")
                
            ElseIf Range("D2") < 60 Then
            
                Sleep 1000
                Range("D2") = Range("D2") + 1
                
            Else
            
                MsgBox "範囲エラー", vbCritical, Title:="ERROR"
                Exit Sub
                
            End If
            
            Count = Count + 1
            
            If Count = 9999999 Then
                Count = 0
                
            Else
            
            End If
            
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
    
        If Range("D2") < 60 Then
        
            Range("D2") = Range("D2") + 1
            
        ElseIf Range("D2") = 60 Then
         
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
