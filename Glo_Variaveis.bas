Attribute VB_Name = "Glo_Variaveis"
Global sStatusMsg As String * 1 '0=ok,1=erro previsto,2-erro imprevisto, msg do sistema
Global sData As String * 10 ' data da ocorrencia
Global sHora As String * 5 'hora da ocorrencia
Global sTipo As String * 29 'tipo da movimentacao a ser gerada
Global sCodFun As String * 5 'codigo do funcionario
Global sMsg As String * 60 'Msg da ocorrencia

Type ArqTexto
     Texto As String * 110
     FFinal As String * 2
End Type
Global sTexto As ArqTexto

Type ArqTextoSql
     TextoSql As String * 11000
     FFinal As String * 2
End Type
Global ArqTextoSql As ArqTextoSql

Public Function retirarComentario(ByVal linha As String) As String
    
    Dim result As String
    Dim posComentario As Integer
    
    posComentario = InStr(1, linha, "//", vbTextCompare)
    If posComentario > 0 Then
        result = Trim(Mid(linha, 1, posComentario - 1))
    Else
        result = linha
    End If
    
    retirarComentario = result
    
End Function

Public Function Retorna_DiaSemana(ByVal dDATA As String) As String

If VBA.Weekday(CDate(dDATA)) = 1 Then Retorna_DiaSemana = "DOMINGO": Exit Function
If VBA.Weekday(CDate(dDATA)) = 2 Then Retorna_DiaSemana = "SEGUNDA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 3 Then Retorna_DiaSemana = "TERCA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 4 Then Retorna_DiaSemana = "QUARTA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 5 Then Retorna_DiaSemana = "QUINTA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 6 Then Retorna_DiaSemana = "SEXTA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 7 Then Retorna_DiaSemana = "SABADO": Exit Function

End Function

Public Function Controle_Dif_Horas(ByVal HoraI As String, ByVal HoraF As String) As String
Dim sHora As String * 2
Dim sMinuto As String * 2
Dim sSegundo As String * 2
Dim nMinuto As Double

Rem verificar se as horas estão preenchidas

If HoraI = "00:00:00" Then
   If HoraF <> "00:00:00" Then
      Controle_Dif_Horas = HoraF: Exit Function
   Else
      Controle_Dif_Horas = "00:00:00": Exit Function
   End If
End If

If HoraF = "00:00:00" Then Controle_Dif_Horas = HoraI: Exit Function

sHora = Format(DateDiff("h", CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO), CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO & ":00")), "00")
If Val(sHora) <> 0 Then
   If Val(sHora) < 0 Then sHora = Format(Val(sHora) * -1, "00")
End If

nMinuto = Format(DateDiff("n", CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_REALIZADO), CDate("01/01/2000 " & Enviar_Email_TresdiasUteis_Ged!TOTAL_SOLICITADO & ":00")), "00")



End Function

Public Function Seg2Tempo(ByVal Segundos As Long) As String
Dim X As Integer, Conta As Long
Dim xDias As Long, xHoras As Long, xMinutos As Long, xSegundos As Long

Conta = Segundos
xDias = Conta \ 86400: Conta = Conta - (xDias * 86400)
xHoras = Conta \ 3600: Conta = Conta - (xHoras * 3600)
xMinutos = Conta \ 60: Conta = Conta - (xMinutos * 60)
xSegundos = Conta

Seg2Tempo$ = IIf(Segundos >= 86400, xDias & IIf(xDias > 1, " dias + ", " dia + "), vbNullString) & IIf(Segundos >= 3600, Format(xHoras, "00") & ":", Format(xHoras, "00") & ":") & Format(xMinutos, "00") & ":" & Format(xSegundos, "00")
End Function
