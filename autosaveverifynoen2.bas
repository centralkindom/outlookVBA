Attribute VB_Name = "autosaveverifynoen"
Private mysuj, mysender, attaname, itemBody3
Dim attaCount4 As Integer
Dim flagifhasatta2 As Boolean
Private mycnt As Integer

Dim mi4 As String





Public Sub SVAttafromnoen(item As Outlook.MailItem)

    Dim myattachment
    
   
        mysuj = item.Subject
        mysender = item.SenderEmailAddress
        itemBody3 = ""
        itemBody3 = item.Body
        
          attaname = ""
          myattachment = ""
          Dim n8 As Integer
             For Each myattachment In item.Attachments
         If myattachment.Size > 0 Then
         
         If myattachment.FileName Like "*.jpg" Or myattachment.FileName Like "*.png" Then
        flagifhasatta2 = True
        
         Else
         n8 = n8 + 1
        attaname = attaname & "<<" & myattachment.FileName & ">> "
        End If
       End If
Next myattachment
attaCount4 = 0
If n8 = 0 Then
attaCount4 = 0
Else
attaCount4 = n8
End If

      If Len(attaname) > 0 Then
      flagifhasatta2 = False
      Else
      flagifhasatta2 = True
      GoTo 120:
      End If
      
      
      
Dim mi2 As Integer

Dim mi3 As Integer

Dim mi5 As Integer


mi3 = Len(mysuj)

mi2 = InStr(1, mysuj, "<ID", vbTextCompare)
mi5 = InStr(1, mysuj, ">", vbTextCompare)
If mi2 > 0 And mi5 > 0 Then
mi4 = ""
mi4 = Mid(mysuj, Int(mi2) + 3, (Int(mi5) - Int(mi2)) - 3)


Else
MsgBox mysuj & Chr(10) & "ID���ƻ����޷������Ӣ��У�鸽������Ҫ�ֹ�����"
mycnt = 0
Exit Sub
End If

    
  
   
   
 Rem �������ֵ�
 Set dt = CreateObject("Scripting.Dictionary")
   dt.Add "pg@ana.us", "EN"
   dt.Add "edfmow@aina.com", "RU"
   dt.Add "manuala@aina.it", "IT"
   dt.Add "nicdfolaeto@ana.it", "IT2"
   dt.Add "huadfngyng@aa.com", "FR"
   dt.Add "togdfashi@aiina.co.jp", "JP"
   dt.Add "doneng@aina.com", "DE"
   dt.Add "ma@airna.es", "ES"
   dt.Add "jerry@airna.com.br", "PO"
   dt.Add "airyxt@ner.com", "KO"
   dt.Add "mddi_chen@sa.cn", "micceshi"
   dt.Add "mien@ana.com", "mieshi2"
   dt.Add "moting@aina.com", "RU2"
   dt.Add "zhgan@ana.com", "JP2"
   
  
   Rem newly add
   Dim addsavefolder
   If dt.exists(mysender) Then
      addsavefolder = dt(mysender) & "_Verify"
      Else
      addsavefolder = "UnknownSender"
      Rem newly add 20161125
   End If
   
    
 Rem �������ֵ�
 
 Set yyy = CreateObject("Scripting.FileSystemObject")
  If yyy.FolderExists("D:\�����ܽ�\201604���빤���ӹ�\" & mi4 & "\��������У�鷵��\" & addsavefolder) = False Then

 
  If Len(attaname) > 0 Then
  On Error GoTo 133
  MkDir "D:\�����ܽ�\20160���빤���ӹ�\" & mi4 & "\��������У�鷵��\" & addsavefolder
  End If
  End If
  
  If Len(attaname) > 0 Then
  On Error GoTo 133
  SaveAttachment item, "D:\�����ܽ�\201604���빤���ӹ�\" & mi4 & "\��������У�鷵��\" & addsavefolder & "\"
   MsgBox mi4 & Chr(10) & "�Զ������Ӣ��У�鸽���ɹ�"
   End If
Rem ����дУ��


Dim mi222 As Integer

Dim mi333 As Integer
Dim mi444 As String
Dim mi555 As Integer
Dim mi666 As Integer
Dim mi777 As Integer

Dim mi888 As String

mi888 = ""
mi333 = Len(mysuj)

mi222 = InStr(1, mysuj, "<ID", vbTextCompare)
mi555 = InStr(1, mysuj, ">", vbTextCompare)
mi444 = Mid(mysuj, Int(mi222) + 3, (Int(mi555) - Int(mi222)) - 3)
mi666 = InStr(10, mi444, "_", vbTextCompare)
mi777 = InStr(mi666 + 1, mi444, "_", vbTextCompare)


mi888 = Mid(mi444, mi666 + 1, (mi777 - mi666) - 1)

If Len(attaname) > 0 Then

Open "D:\�����ܽ�\20160429���빤���ӹ�\" & mi4 & "\log.txt" For Append As #1

Write #1, "�����Ӣ��У��", mi888, mysender, attaname, Now()





Close #1
   

Open "D:\�����ܽ�\20160429���빤���ӹ�\" & mi4 & "\��������У�鷵��\log.txt" For Append As #23

Write #23, "�����Ӣ��У��", mi888, mysender, attaname, Now()
Call У����ս������noEN

Close #23
   
End If

   



Rem ���������
120:
Dim miend
miend = 20
Rem miend = InStr(1, item.Body, "��", vbTextCompare)
If miend < 0 Or miend = 0 Or miend > 50 Then
miend = InStr(1, item.Body, "_", vbTextCompare)
ElseIf miend < 0 Or miend = 0 Or miend > 50 Then
miend = InStr(1, item.Body, "<", vbTextCompare)
ElseIf miend < 0 Or miend = 0 Or miend > 50 Then
miend = InStr(1, item.Body, "-", vbTextCompare)
ElseIf miend < 0 Or miend = 0 Or miend > 50 Then
miend = 20
Else
miend = 20
End If


Dim mi20 As Integer

Dim mi30 As Integer
Dim mi40 As String
Dim mi50 As Integer


mi30 = Len(mysuj)

mi20 = InStr(1, mysuj, "<ID", vbTextCompare)
mi50 = InStr(1, mysuj, ">", vbTextCompare)
If mi20 > 0 And mi50 > 0 Then
mi40 = Mid(mysuj, Int(mi20) + 3, (Int(mi50) - Int(mi20)) - 3)
End If



If flagifhasatta2 = True Then
Open "D:\�����ܽ�\20160429���빤���ӹ�\" & mi40 & "\log.txt" For Append As #41



Write #41, "��Ӣ��У�鷵�ص���û�и��������忴�ʼ�", mi888, mysender, Now(), Mid(item.Body, 1, miend)
Call У����ս������noEN

Close #41
 
On Error GoTo 134
 



Open "D:\�����ܽ�\20160429���빤���ӹ�\" & mi40 & "\��������У�鷵��\log.txt" For Append As #43



Write #43, "��Ӣ��У�鷵�ص���û�и��������忴�ʼ�", mi888, mysender, Now(), Mid(item.Body, 1, miend)


Close #43
mycnt = 0
Exit Sub
End If


Rem ���������




 Set myattachment = Nothing
    Set item = Nothing
   
      
  mysender = ""
   attaname = ""
   mycnt = 0
Exit Sub

133:
MsgBox "û���ҵ������ļ��У���Ҫ�ֶ������Ӣ��У�鷵�ظ���"
134:
MsgBox "subjectû��ID"


mycnt = 0
End Sub

' ���渽��
' pathΪ����·����conditionΪ������ƥ������
Private Sub SaveAttachment(ByVal item As Object, path$)
    Dim olAtt As Attachment
    Dim i As Integer
    Dim mflag As Boolean
    Dim mn As Integer
    
    
    If item.Attachments.Count > 0 Then
        For i = 1 To item.Attachments.Count
            Set olAtt = item.Attachments(i)
            
            ' save the attachment
            If olAtt.FileName Like "*.docx" Then
                      
            
                olAtt.SaveAsFile path & olAtt.FileName
                attnewname = attnewname & "," & olAtt.FileName
                mflag = True
                mn = mn + 1
                
             ElseIf olAtt.FileName Like "*.doc" Then
               olAtt.SaveAsFile path & olAtt.FileName
               attnewname = attnewname & "," & olAtt.FileName
               mflag = True
               mn = mn + 1
               
               
             
              ElseIf olAtt.FileName Like "*.xlsx" Then
               olAtt.SaveAsFile path & olAtt.FileName
                attnewname = attnewname & "," & olAtt.FileName
              mflag = True
              mn = mn + 1
                
               
              ElseIf olAtt.FileName Like "*.xls" Then
               olAtt.SaveAsFile path & olAtt.FileName
                 attnewname = attnewname & "," & olAtt.FileName
                 mflag = True
                 mn = mn + 1
                 
                
              
              ElseIf olAtt.FileName Like "*.xlsm" Then
               olAtt.SaveAsFile path & olAtt.FileName
                 attnewname = attnewname & "," & olAtt.FileName
                 mflag = True
                 mn = mn + 1
                
           
             
               ElseIf olAtt.FileName Like "*.txt" Then
               
               olAtt.SaveAsFile path & olAtt.FileName
               attnewname = attnewname & "," & olAtt.FileName
                 mflag = True
                 mn = mn + 1
               
            
               ElseIf olAtt.FileName Like "*.ppt" Then
               olAtt.SaveAsFile path & olAtt.FileName
                attnewname = attnewname & "," & olAtt.FileName
                 mflag = True
                 mn = mn + 1
                
            
                ElseIf olAtt.FileName Like "*.pptx" Then
               olAtt.SaveAsFile path & olAtt.FileName
                  attnewname = attnewname & "," & olAtt.FileName
                 mflag = True
                 mn = mn + 1
                
            
              ElseIf olAtt.FileName Like "*.csv" Then
               olAtt.SaveAsFile path & olAtt.FileName
             attnewname = attnewname & "," & olAtt.FileName
                  mflag = True
                  mn = mn + 1
                 
           
              ElseIf olAtt.FileName Like "*.rtf" Then
               olAtt.SaveAsFile path & olAtt.FileName
                 attnewname = attnewname & "," & olAtt.FileName
              mflag = True
              mn = mn + 1
              
             
                ElseIf olAtt.FileName Like "*.pdf" Then
               olAtt.SaveAsFile path & olAtt.FileName
              attnewname = attnewname & "," & olAtt.FileName
             mflag = True
             mn = mn + 1
             
                
                 ElseIf olAtt.FileName Like "*.rar" Then
               olAtt.SaveAsFile path & olAtt.FileName
              attnewname = attnewname & "," & olAtt.FileName
            mflag = True
            mn = mn + 1
            
              ElseIf olAtt.FileName Like "*.zip" Then
               olAtt.SaveAsFile path & olAtt.FileName
              attnewname = attnewname & "," & olAtt.FileName
           mflag = True
           mn = mn + 1
            
                 ElseIf olAtt.FileName Like "*.html" Then
               olAtt.SaveAsFile path & olAtt.FileName
              attnewname = attnewname & "," & olAtt.FileName
            mflag = True
            mn = mn + 1
             
                   ElseIf olAtt.FileName Like "*.htm" Then
               olAtt.SaveAsFile path & olAtt.FileName
              attnewname = attnewname & "," & olAtt.FileName
             mflag = True
             mn = mn + 1
            
             Else
             
             End If
             
             
        Next
    End If
   
    Set olAtt = Nothing
   
    Set myattachment = Nothing
    Set item = Nothing
     
    
End Sub


Sub У����ս������noEN()

Dim mi2 As Integer

Dim mi3 As Integer
Dim mi4 As String
Dim mi5 As Integer


mi3 = Len(mysuj)



mi2 = InStr(1, mysuj, "<ID", vbBinaryCompare)
mi5 = InStr(1, mysuj, ">", vbBinaryCompare)
If mi2 > 0 And mi5 > 0 Then
mi4 = Mid(mysuj, Int(mi2) + 3, (Int(mi5) - Int(mi2)) - 3)
Else
MsgBox "Subjectû��ID"
Exit Sub
End If



Dim mi222 As Integer

Dim mi333 As Integer
Dim mi444 As String
Dim mi555 As Integer
Dim mi666 As Integer
Dim mi777 As Integer

Dim mi888 As String


mi333 = Len(mysuj)

mi222 = InStr(1, mysuj, "<ID", vbBinaryCompare)
mi555 = InStr(1, mysuj, ">", vbBinaryCompare)
mi444 = Mid(mysuj, Int(mi222) + 3, (Int(mi555) - Int(mi222)) - 3)
mi666 = InStr(10, mi444, "_", vbBinaryCompare)
mi777 = InStr(mi666 + 1, mi444, "_", vbBinaryCompare)



mi888 = Mid(mi444, mi666 + 1, (mi777 - mi666) - 1)

Open "D:\�����ܽ�\20160429���빤���ӹ�\" & mi4 & "\ReceiveMailBonusLog.txt" For Append As #42

Write #42, mi888, "���鷵�ؽ���", mysender, "��������", attaname, attaCount4, Now()

Close #42
   
Open "D:\�����ܽ�\20160429���빤���ӹ�\���⽱�����\ReceiveMailBonusLog.txt" For Append As #24

Write #24, mi888, "���鷵�ؽ���", mysender, "���ս�������", attaname, attaCount4, Now(), Mid(itemBody3, 1, 100)

Close #24


mycnt = mycnt + 1
If itemBody3 Like "*leave*" Or itemBody3 Like "*vocation*" Or itemBody3 Like "*Leave*" Or itemBody3 Like "*Vocation*" Then
MsgBox "У����Ա�ݼ�"
Else
��������д��EXCELb mi888, "���鷵�ؽ���", mysender, "���ս�������", attaname, attaCount4, Now(), Mid(itemBody3, 1, 200)
End If
End Sub

Rem ��������д��EXCELb mi888, "���鷵�ؽ���", mysender, "���ս�������", attaname, attaCount4, Now(), Mid(itemBody3, 1, 300)
Sub ��������д��EXCELb(a, b, c, d, e, f, g, h)
Set Conn = CreateObject("adodb.connection")
Set rst = CreateObject("ADODB.recordset")
Conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;extended properties=Excel 12.0;data source=" & "D:\�����ܽ�\20160429���빤���ӹ�\���⽱�����" & "/�����������ݿ�.xls"
rst.Open "select *  from [����$]", Conn, , adLockOptimistic
rst.addnew
rst.fields("����") = CDate(Format(Now(), yyyy - mm - dd))
rst.fields("��Ŀ����") = Mid(a, 1, 200)
rst.fields("����") = b
rst.fields("У�鷵�巢����") = Mid(c, 1, 200)
rst.fields("������ʶ") = d
rst.fields("�������Ը�������") = Mid(e, 1, 200)
rst.fields("�������Ը�����") = CInt(f)
rst.fields("ʱ���") = g
rst.fields("�ʼ�����") = h
rst.fields("�ʼ���") = CInt(1)

rst.Update
rst.Close
Conn.Close
Set rst = Nothing
Set Conn = Nothing

If (mycnt <= 1) Then
MsgBox "�����뵽���ݿ�"
End If

End Sub



