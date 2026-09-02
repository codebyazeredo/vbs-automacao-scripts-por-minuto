Option Explicit

Dim FSO
Dim WshShell
Dim pastaBase
Dim pastaBat
Dim pastaLog
Dim arquivoConfigPath
Dim arquivoEstado
Dim arquivoConfigHandle
Dim arquivoSaida
Dim logFile
Dim estado
Dim linha
Dim partes
Dim nomeBat
Dim intervalo
Dim caminhoBat
Dim caminhoSaida
Dim caminhoLog
Dim caminhoLogErro
Dim nomeLog
Dim agora
Dim ultimaExecucao
Dim deveExecutar
Dim inicio
Dim fim
Dim exitCode
Dim houveErro
Dim executouAlgumBat
Dim mensagemSaida
Dim mensagemErro
Dim comando
Dim aspas

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

aspas = Chr(34)

pastaBase = FSO.GetParentFolderName(WScript.ScriptFullName)
pastaBat = pastaBase & "\bat"
pastaLog = pastaBase & "\logs"

arquivoConfigPath = pastaBase & "\config.ini"
arquivoEstado = pastaBase & "\estado.ini"

If Not FSO.FolderExists(pastaBat) Then
    WScript.Echo "ERRO: Pasta bat nao encontrada:"
    WScript.Echo pastaBat
    WScript.Quit 1
End If

If Not FSO.FolderExists(pastaLog) Then
    FSO.CreateFolder pastaLog
End If

If Not FSO.FileExists(arquivoConfigPath) Then
    WScript.Echo "ERRO: config.ini nao encontrado:"
    WScript.Echo arquivoConfigPath
    WScript.Quit 1
End If

Set estado = CarregarEstado()

houveErro = False
executouAlgumBat = False

nomeLog = "execucao_" & _
          Year(Now) & "-" & _
          Right("0" & Month(Now), 2) & "-" & _
          Right("0" & Day(Now), 2) & "_" & _
          Right("0" & Hour(Now), 2) & "-" & _
          Right("0" & Minute(Now), 2) & "-" & _
          Right("0" & Second(Now), 2) & _
          ".log"

caminhoLog = pastaLog & "\" & nomeLog
caminhoLogErro = pastaLog & "\" & Left(nomeLog, Len(nomeLog) - 4) & "_ERRO.log"

Set logFile = FSO.OpenTextFile(caminhoLog, 8, True)

logFile.WriteLine "INICIO DA EXECUCAO"
logFile.WriteLine ""
logFile.WriteLine "============================================================="
logFile.WriteLine ""

Set arquivoConfigHandle = FSO.OpenTextFile(arquivoConfigPath, 1, False)

Do Until arquivoConfigHandle.AtEndOfStream
    linha = Trim(arquivoConfigHandle.ReadLine)
    If linha <> "" Then
        If Left(linha, 1) <> "#" And Left(linha, 1) <> ";" Then
            partes = Split(linha, "=")
            If UBound(partes) >= 1 Then
                nomeBat = Trim(partes(0))
                intervalo = Trim(partes(1))
                If nomeBat <> "" And IsNumeric(intervalo) Then
                    intervalo = CLng(intervalo)
                    If intervalo > 0 Then
                        caminhoBat = pastaBat & "\" & nomeBat
                        If FSO.FileExists(caminhoBat) Then
                            agora = AgoraUnix()
                            deveExecutar = False
                            If estado.Exists(nomeBat) Then
                                If IsNumeric(estado(nomeBat)) Then
                                    ultimaExecucao = CLng(estado(nomeBat))
                                    If (agora - ultimaExecucao) >= (intervalo * 60) Then
                                        deveExecutar = True
                                    End If
                                Else
                                    deveExecutar = True
                                End If
                            Else
                                deveExecutar = True
                            End If
                            If deveExecutar Then
                                executouAlgumBat = True
                                inicio = Now
                                caminhoSaida = pastaLog & "\_" & Replace(nomeBat, ".bat", "") & "_" & AgoraUnix() & ".tmp"
                                comando = "cmd.exe /d /c " & aspas & "cd /d " & aspas & pastaBat & aspas & " && call " & aspas & caminhoBat & aspas & " > " & aspas & caminhoSaida & aspas & " 2>&1" & aspas
                                exitCode = WshShell.Run(comando, 0, True)
                                fim = Now
                                mensagemSaida = ""
                                If FSO.FileExists(caminhoSaida) Then
                                    Set arquivoSaida = FSO.OpenTextFile(caminhoSaida, 1, False)
                                    Do Until arquivoSaida.AtEndOfStream
                                        linha = arquivoSaida.ReadLine
                                        If Trim(linha) <> "" Then
                                            mensagemSaida = mensagemSaida & linha & vbCrLf
                                        End If
                                    Loop
                                    arquivoSaida.Close
                                    Set arquivoSaida = Nothing
                                End If
                                If exitCode = 0 Then
                                    logFile.WriteLine "[" & FormatDateTime(fim, 3) & "] " & nomeBat & " | SUCCESS | SAIDA: 0 | DURACAO: " & DateDiff("s", inicio, fim) & " seg"
                                    estado(nomeBat) = CStr(AgoraUnix())
                                Else
                                    houveErro = True
                                    mensagemErro = Trim(mensagemSaida)
                                    logFile.WriteLine "[" & FormatDateTime(fim, 3) & "] " & nomeBat & " | ERROR | SAIDA: " & exitCode & " | DURACAO: " & DateDiff("s", inicio, fim) & " seg"
                                    If mensagemErro <> "" Then
                                        logFile.WriteLine "    ERRO: " & Replace(mensagemErro, vbCrLf, " | ")
                                    Else
                                        logFile.WriteLine "    ERRO: Nenhuma mensagem retornada pelo BAT."
                                    End If
                                End If
                                If FSO.FileExists(caminhoSaida) Then
                                    FSO.DeleteFile caminhoSaida, True
                                End If
                            End If
                        Else
                            houveErro = True
                            logFile.WriteLine "[" & FormatDateTime(Now, 3) & "] " & nomeBat & " | ERROR | BAT NAO ENCONTRADO"
                        End If
                    Else
                        houveErro = True
                        logFile.WriteLine "[" & FormatDateTime(Now, 3) & "] " & nomeBat & " | ERROR | INTERVALO INVALIDO: " & intervalo
                    End If
                Else
                    houveErro = True
                    logFile.WriteLine "[" & FormatDateTime(Now, 3) & "] ERROR | CONFIGURACAO INVALIDA: " & linha
                End If
            Else
                houveErro = True
                logFile.WriteLine "[" & FormatDateTime(Now, 3) & "] ERROR | LINHA INVALIDA: " & linha
            End If
        End If
    End If
Loop

arquivoConfigHandle.Close
Set arquivoConfigHandle = Nothing

SalvarEstado estado, arquivoEstado

If Not executouAlgumBat Then
    logFile.WriteLine ""
    logFile.WriteLine "[" & FormatDateTime(Now, 3) & "] NENHUM BAT PENDENTE PARA EXECUCAO"
End If

logFile.WriteLine ""
logFile.WriteLine "============================================================="
logFile.WriteLine ""
logFile.WriteLine "[" & FormatDateTime(Now, 3) & "] FIM DA EXECUCAO"

logFile.Close
Set logFile = Nothing

If houveErro Then
    If FSO.FileExists(caminhoLog) Then
        If FSO.FileExists(caminhoLogErro) Then
            FSO.DeleteFile caminhoLogErro, True
        End If
        FSO.MoveFile caminhoLog, caminhoLogErro
    End If
End If

Set estado = Nothing
Set WshShell = Nothing
Set FSO = Nothing

WScript.Quit 0

Function AgoraUnix()
    AgoraUnix = DateDiff("s", DateSerial(1970, 1, 1), Now)
End Function

Function CarregarEstado()
    Dim dict
    Dim arquivo
    Dim linha
    Dim partes
    Set dict = CreateObject("Scripting.Dictionary")
    If Not FSO.FileExists(arquivoEstado) Then
        Set CarregarEstado = dict
        Exit Function
    End If
    Set arquivo = FSO.OpenTextFile(arquivoEstado, 1, False)
    Do Until arquivo.AtEndOfStream
        linha = Trim(arquivo.ReadLine)
        If linha <> "" Then
            partes = Split(linha, "=")
            If UBound(partes) >= 1 Then
                If IsNumeric(Trim(partes(1))) Then
                    dict(Trim(partes(0))) = Trim(partes(1))
                End If
            End If
        End If
    Loop
    arquivo.Close
    Set arquivo = Nothing
    Set CarregarEstado = dict
End Function

Sub SalvarEstado(dict, caminho)
    Dim arquivo
    Dim chave
    Set arquivo = FSO.CreateTextFile(caminho, True)
    For Each chave In dict.Keys
        arquivo.WriteLine chave & "=" & dict(chave)
    Next
    arquivo.Close
    Set arquivo = Nothing
End Sub