Attribute VB_Name = "ExcelTimer_main"
#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Dim Count As Long 'Timer���[�v�̊Ǘ�
Dim Flashing As Byte
Dim StopSignal As Byte '0�Ȃ�Timer��~�A1�Ȃ烊�Z�b�g
Dim Button As Byte '1�ɂȂ��Timer���[�v����~����
Dim ResetSignal As Byte
Dim CountProhibition As Byte '�蓮�J�E���g�̋֎~

Sub Timer()

    Button = 0
    StopSignal = 0
    CountProhibition = 1
    
    For Count = 0 To 9999999
    
        DoEvents
            
            If Button = 1 Then
                Range("J1") = ""
                StopSignal = 1
                CountProhibition = 0
                Exit Sub
                
            End If
            
            Range("J1") = "�J�E���g���E�E�E"
            
            If Range("D2") < 60 Then
            
                Sleep 1000
                
                Range("D2") = Range("D2") + 1
                
                Range("J1") = ""
                
            ElseIf Range("D2") = 60 Then
                
                Range("B2") = Range("B2") + 1
                Range("D2") = Range("D2") - Range("D2")
                
            Else
            
                MsgBox "�͈̓G���[", vbCritical, Title:="ERROR"
                
            End If
            
            Count = Count + 1
            
            If Count = 9999999 Then
                Count = 0
                
            Else
            
            End If
            
    Next Count
    
End Sub

Sub eButton()

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

    If CountProhibition = 0 Then
    
        If Range("D2") < 59 Then
        
            Range("D2") = Range("D2") + 1
            
            
        ElseIf Range("D2") = 59 Then
         
            Range("B2") = Range("B2") + 1
            Range("D2") = Range("D2") - Range("D2")
            
        Else
        
           MsgBox "�͈̓G���[", vbCritical, Title:="ERROR"
           
        End If
    
        StopSignal = 1
    
    Else
    
    End If
    
End Sub

Sub Minus_s()

    If CountProhibition = 0 Then

        If Range("D2") = 0 And Range("B2") = 0 Then
        
        ElseIf Range("D2") = 0 And Range("B2") >= 1 Then
            
            Range("B2") = Range("B2") - 1
            Range("D2") = 59
            
        ElseIf 1 <= Range("D2") And Range("D2") < 60 Then
        
            Range("D2") = Range("D2") - 1
        
        Else
            MsgBox "�͈̓G���[", vbCritical, Title:="ERROR"
        
        End If
    
        StopSignal = 1
    
    Else
         
    End If
    
End Sub
