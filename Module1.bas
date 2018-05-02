Attribute VB_Name = "Module1"
Option Explicit

Public Function FechaActual() As String
    FechaActual = Format(Date, "yyyy-MM-dd")
End Function

Public Function MailTo(ByVal Dest As String, ByVal Subject As String, ByVal Body As String, Optional ByVal Adjunto As Integer) As Boolean
  On Error GoTo Err_MailTo
  Dim objOLMAIL As Object
  Dim objMAIL As Object
  Dim objATCH As Object
  Dim objDEST As Object
  Dim Cont, Max As Byte
  Dim correos() As String
  Cont = 0
  Max = 0
  Dim ValidadorInterfase As String
  ValidadorInterfase = "X:\VALIDA.CCL.TXT"

  Set objOLMAIL = CreateObject("Outlook.Application")
  Set objMAIL = objOLMAIL.CreateItem(0)
  objMAIL.Subject = Subject
  objMAIL.Body = Body

  Set objDEST = objMAIL.Recipients
  'preramose el tema para tener mas de un contacto
  Dest = Dest & ";"
  correos = Split(Dest, ";")
  Max = UBound(correos)
  While Cont < Max
    objDEST.Add correos(Cont)
    Cont = Cont + 1
  Wend
 
  If Not objDEST.ResolveAll Then
       MsgBox "Error en destinatario de correo " & Dest & "."
       GoTo Exit_MailTo
  End If

  If Adjunto > 0 Then
    Cont = 0
     Set objATCH = objMAIL.Attachments
        If Adjunto > 0 Then
            objATCH.Add ValidadorInterfase
        End If
     'objATCH.Add Adjunto
  End If

  objMAIL.Send

  MailTo = True

Exit_MailTo:
  Set objDEST = Nothing
  Set objATCH = Nothing
  Set objMAIL = Nothing
  Set objOLMAIL = Nothing
  Exit Function

Err_MailTo:
  MsgBox Err.Description
  Resume Exit_MailTo

End Function
