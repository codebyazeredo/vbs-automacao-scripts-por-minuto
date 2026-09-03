Option Explicit

Const DIAS_RETENCAO_LOGS = 60

Dim FSO
Dim WshShell
Dim pastaBase
Dim pastaBat
Dim pastaLog
Dim pastaLock
Dim arquivoConfigPath
Dim arquivoEstado
Dim arquivoConfigHandle
Dim arquivoSaida
Dim estado
Dim linha
Dim partes
Dim nomeBat
Dim intervalo
Dim caminhoBat
Dim caminhoSaida
Dim caminhoLogDiario
Dim caminhoLogErroDiario
Dim nomeLogDiario
Dim nomeLogErroDiario
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
Dim logBufferNormal
Dim logBufferErro
Dim caminhoLock
Dim podeExecutar

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

aspas = Chr(34)
logBufferNormal = ""
logBufferErro = ""

pastaBase = FSO.GetParentFolderName(WScript.ScriptFullName)
pastaBat = pastaBase & "\bat"
pastaLog = pastaBase & "\logs"
pastaLock = pastaBase & "\locks"

arquivoConfigPath = pastaBase & "\config.ini"
arquivoEstado = pastaBase & "\estado.ini"

If Not FSO.FolderExists(pastaBat) Then
    WScript.Echo "ERRO: Pasta bat nao encontrada: " & pastaBat
    WScript.Quit 1
End If

If Not FSO.FolderExists(pastaLog) Then FSO.CreateFolder pastaLog
If Not FSO.FolderExists(pastaLock) Then FSO.CreateFolder pastaLock

If Not FSO.FileExists(arquivoConfigPath) Then
    WScript.Echo "ERRO: config.ini nao encontrado: " & arquivoConfigPath
    WScript.Quit 1
End If

LimparLogsAntigos

Set estado = CarregarEstado()

houveErro = False
executouAlgumBat = False

nomeLogDiario = "execucao_" & Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & ".log"
nomeLogErroDiario = "erros_" & Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & ".log"

caminhoLogDiario = pastaLog & "\" & nomeLogDiario
caminhoLogErroDiario = pastaLog & "\" & nomeLogErroDiario

Log "================================================================"
Log "INÍCIO DA EXECUÇÃO: " & FormatDateTime(Now, 0)
Log "================================================================"

LogError "================================================================"
LogError "INÍCIO DOS REGISTROS DE ERRO: " & FormatDateTime(Now, 0)
LogError "================================================================"

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
                                caminhoLock = pastaLock & "\" & nomeBat & ".lock"
                                podeExecutar = CriarLock(caminhoLock)
                                
                                If podeExecutar Then
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
                                        Log "[" & FormatDateTime(fim, 3) & "] " & nomeBat
                                        Log "  Status: OK (Saída: 0) | Duração: " & DateDiff("s", inicio, fim) & "s"
                                        Log ""
                                        
                                        estado(nomeBat) = CStr(AgoraUnix())
                                    Else
                                        houveErro = True
                                        mensagemErro = Trim(mensagemSaida)
                                        
                                        LogError "[" & FormatDateTime(fim, 3) & "] " & nomeBat
                                        LogError "  Status: ERRO (Saída: " & exitCode & ") | Duração: " & DateDiff("s", inicio, fim) & "s"
                                        
                                        If mensagemErro <> "" Then
                                            LogError "  Detalhes: " & Replace(mensagemErro, vbCrLf, " ")
                                        Else
                                            LogError "  Detalhes: Nenhuma mensagem retornada pelo BAT."
                                        End If
                                        LogError ""
                                    End If
                                    
                                    If FSO.FileExists(caminhoSaida) Then
                                        FSO.DeleteFile caminhoSaida, True
                                    End If
                                    
                                    RemoverLock caminhoLock
                                Else
                                    Log "[" & FormatDateTime(Now, 3) & "] " & nomeBat & " | IGNORADO (Já em execução)"
                                    Log ""
                                End If
                            End If
                        Else
                            houveErro = True
                            LogError "[" & FormatDateTime(Now, 3) & "] " & nomeBat
                            LogError "  Status: ERRO | Arquivo BAT não encontrado em: " & caminhoBat
                            LogError ""
                        End If
                    Else
                        houveErro = True
                        LogError "[" & FormatDateTime(Now, 3) & "] " & nomeBat
                        LogError "  Status: ERRO | Intervalo de tempo inválido: " & intervalo
                        LogError ""
                    End If
                Else
                    houveErro = True
                    LogError "[" & FormatDateTime(Now, 3) & "] Configuração inválida no arquivo INI: " & linha
                    LogError ""
                End If
            Else
                houveErro = True
                LogError "[" & FormatDateTime(Now, 3) & "] Linha mal formatada no arquivo INI: " & linha
                LogError ""
            End If
        End If
    End If
Loop

arquivoConfigHandle.Close
Set arquivoConfigHandle = Nothing

SalvarEstado estado, arquivoEstado

If Not executouAlgumBat Then
    Log "[" & FormatDateTime(Now, 3) & "] NENHUM BAT PENDENTE PARA EXECUÇÃO"
    Log ""
End If

Log "================================================================"
Log "FIM DA EXECUÇÃO: " & FormatDateTime(Now, 3)
Log "================================================================"
Log ""

LogError "================================================================"
LogError "FIM DOS REGISTROS DE ERRO: " & FormatDateTime(Now, 0)
LogError "================================================================"
LogError ""

If Not houveErro Then
    GravarNoArquivo caminhoLogDiario, logBufferNormal
Else
    GravarNoArquivo caminhoLogErroDiario, logBufferErro
End If

Set estado = Nothing
Set WshShell = Nothing
Set FSO = Nothing

WScript.Quit 0

Sub Log(texto)
    logBufferNormal = logBufferNormal & texto & vbCrLf
End Sub

Sub LogError(texto)
    logBufferErro = logBufferErro & texto & vbCrLf
End Sub

Sub GravarNoArquivo(caminho, conteudo)
    Dim arquivo
    Set arquivo = FSO.OpenTextFile(caminho, 8, True)
    arquivo.Write conteudo
    arquivo.Close
    Set arquivo = Nothing
End Sub

Sub LimparLogsAntigos()
    Dim arquivo
    Dim limite
    limite = DateAdd("d", - DIAS_RETENCAO_LOGS, Now)
    
    For Each arquivo In FSO.GetFolder(pastaLog).Files
        If LCase(FSO.GetExtensionName(arquivo.Name)) = "log" Then
            If arquivo.DateLastModified < limite Then
                FSO.DeleteFile arquivo.Path, True
            End If
        End If
    Next
End Sub

Function AgoraUnix()
    AgoraUnix = DateDiff("s", DateSerial(1970, 1, 1), Now)
End Function

Function CarregarEstado()
    Dim dict, arquivo, linha, partes
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
    Dim arquivo, chave
    Set arquivo = FSO.CreateTextFile(caminho, True)
    For Each chave In dict.Keys
        arquivo.WriteLine chave & "=" & dict(chave)
    Next
    arquivo.Close
    Set arquivo = Nothing
End Sub

Function CriarLock(caminho)
    Dim arquivo
    CriarLock = False
    On Error Resume Next
    Set arquivo = FSO.CreateTextFile(caminho, False)
    If Err.Number = 0 Then
        arquivo.WriteLine "INICIO=" & FormatDateTime(Now, 0)
        arquivo.Close
        Set arquivo = Nothing
        CriarLock = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Sub RemoverLock(caminho)
    On Error Resume Next
    If FSO.FileExists(caminho) Then
        FSO.DeleteFile caminho, True
    End If
    On Error GoTo 0
End Sub