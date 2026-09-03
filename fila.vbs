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
Dim logBuffer
Dim caminhoLock
Dim podeExecutar

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

aspas = Chr(34)
logBuffer = ""

pastaBase = FSO.GetParentFolderName(WScript.ScriptFullName)
pastaBat = pastaBase & "\bat"
pastaLog = pastaBase & "\logs"
pastaLock = pastaBase & "\locks"

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

If Not FSO.FolderExists(pastaLock) Then
    FSO.CreateFolder pastaLock
End If

If Not FSO.FileExists(arquivoConfigPath) Then
    WScript.Echo "ERRO: config.ini nao encontrado:"
    WScript.Echo arquivoConfigPath
    WScript.Quit 1
End If

LimparLogsAntigos

Set estado = CarregarEstado()

houveErro = False
executouAlgumBat = False

nomeLogDiario = "execucao_" & _
    Year(Now) & "-" & _
    Right("0" & Month(Now), 2) & "-" & _
    Right("0" & Day(Now), 2) & ".log"

nomeLogErroDiario = "erros_" & _
    Year(Now) & "-" & _
    Right("0" & Month(Now), 2) & "-" & _
    Right("0" & Day(Now), 2) & ".log"

caminhoLogDiario = pastaLog & "\" & nomeLogDiario
caminhoLogErroDiario = pastaLog & "\" & nomeLogErroDiario

Log ""
Log "###############################################################"
Log "INICIO DA EXECUCAO - " & FormatDateTime(Now, 0)
Log ""

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
                                
                                ' Impede duas execuções simultâneas do mesmo BAT.
                                podeExecutar = CriarLock(caminhoLock)
                                
                                If podeExecutar Then
                                    
                                    executouAlgumBat = True
                                    inicio = Now
                                    
                                    caminhoSaida = pastaLog & "\_" & _
                                        Replace(nomeBat, ".bat", "") & "_" & _
                                        AgoraUnix() & ".tmp"
                                    
                                    comando = "cmd.exe /d /c " & _
                                        aspas & _
                                        "cd /d " & aspas & pastaBat & aspas & _
                                        " && call " & aspas & caminhoBat & aspas & _
                                        " > " & aspas & caminhoSaida & aspas & _
                                        " 2>&1" & _
                                        aspas
                                    
                                    ' Aguarda o BAT terminar antes de continuar.
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
                                        
                                        Log "[" & FormatDateTime(fim, 3) & "] " & _
                                            nomeBat & _
                                            " | SUCCESS | SAIDA: 0 | DURACAO: " & _
                                            DateDiff("s", inicio, fim) & " seg"
                                        
                                        estado(nomeBat) = CStr(AgoraUnix())
                                        
                                    Else
                                        
                                        houveErro = True
                                        
                                        mensagemErro = Trim(mensagemSaida)
                                        
                                        Log "[" & FormatDateTime(fim, 3) & "] " & _
                                            nomeBat & _
                                            " | ERROR | SAIDA: " & _
                                            exitCode & _
                                            " | DURACAO: " & _
                                            DateDiff("s", inicio, fim) & " seg"
                                        
                                        If mensagemErro <> "" Then
                                            
                                            Log "    ERRO: " & _
                                                Replace(mensagemErro, vbCrLf, " | ")
                                            
                                        Else
                                            
                                            Log "    ERRO: Nenhuma mensagem retornada pelo BAT."
                                            
                                        End If
                                        
                                    End If
                                    
                                    If FSO.FileExists(caminhoSaida) Then
                                        FSO.DeleteFile caminhoSaida, True
                                    End If
                                    
                                    ' Libera o BAT para uma nova execução.
                                    RemoverLock caminhoLock
                                    
                                Else
                                    
                                    Log "[" & FormatDateTime(Now, 3) & "] " & _
                                        nomeBat & _
                                        " | IGNORADO | BAT JA ESTA EM EXECUCAO"
                                    
                                End If
                                
                            End If
                            
                        Else
                            
                            houveErro = True
                            
                            Log "[" & FormatDateTime(Now, 3) & "] " & _
                                nomeBat & _
                                " | ERROR | BAT NAO ENCONTRADO"
                            
                        End If
                        
                    Else
                        
                        houveErro = True
                        
                        Log "[" & FormatDateTime(Now, 3) & "] " & _
                            nomeBat & _
                            " | ERROR | INTERVALO INVALIDO: " & _
                            intervalo
                        
                    End If
                    
                Else
                    
                    houveErro = True
                    
                    Log "[" & FormatDateTime(Now, 3) & "] ERROR | " & _
                        "CONFIGURACAO INVALIDA: " & linha
                    
                End If
                
            Else
                
                houveErro = True
                
                Log "[" & FormatDateTime(Now, 3) & "] ERROR | " & _
                    "LINHA INVALIDA: " & linha
                
            End If
            
        End If
        
    End If
    
Loop

arquivoConfigHandle.Close
Set arquivoConfigHandle = Nothing

SalvarEstado estado, arquivoEstado

If Not executouAlgumBat Then
    
    Log "[" & FormatDateTime(Now, 3) & "] " & _
        "NENHUM BAT PENDENTE PARA EXECUCAO"
    
End If

Log ""
Log "FIM DA EXECUCAO - " & FormatDateTime(Now, 3)
Log "###############################################################"
Log ""
Log ""

GravarNoArquivo caminhoLogDiario, logBuffer

If houveErro Then
    GravarNoArquivo caminhoLogErroDiario, logBuffer
End If

Set estado = Nothing
Set WshShell = Nothing
Set FSO = Nothing

WScript.Quit 0


Sub Log(texto)
    
    logBuffer = logBuffer & texto & vbCrLf
    
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
    
    AgoraUnix = DateDiff( _
        "s", _
        DateSerial(1970, 1, 1), _
        Now _
        )
    
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


Function CriarLock(caminho)
    
    Dim arquivo
    
    CriarLock = False
    
    On Error Resume Next
    
    ' CreateTextFile(False) falha se o arquivo ja existir.
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