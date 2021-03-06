'==========================================================================='
'                                                                           '
'  Script de Logon VBS                                                      '
'  Microsoft Windows Server 2016                                            '
'                                                                           '
'==========================================================================='


'==========================================================================='
' Impedindo a Exibição de Erro para o Usuário                               '
'==========================================================================='

On error Resume Next
Err.clear 0


'==========================================================================='
' Sincroniza o Horário da Estação com o Servidor                            '         
'==========================================================================='

'set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") '
'set objShell = CreateObject("WScript.shell") '
'strCmd = "net time \\nomeserver /set /yes" '
'set objExec = objShell.exec(strCmd) '


'==========================================================================='
' Mapear Pastas de acordo com o Grupo do USER                               '
'==========================================================================='

set objNetwork = CreateObject("WScript.Network")
strDom = objNetwork.UserDomain
strUser = objNetwork.UserName
set objUser = GetObject("WinNT://" & strDom & "/" & strUser & ",user")
set FSODrive = CreateObject("Scripting.FileSystemObject")

For Each objGroup In objUser.Groups

  Select Case objGroup.Name
    Case "DL_Setor_Administrativo"
      If not FSODrive.DriveExists("S:") Then
        objNetwork.MapNetworkDrive "S:", "\\SRVHOMOLOGDC1\Adm","true"
      End If

    Case "DL_Enfermeiros"
      If not FSODrive.DriveExists("S:") Then
        objNetwork.MapNetworkDrive "S:", "\\SRVHOMOLOGDC1\Enfermeiros","true"
      End If

    Case "DL_Medicos"
      If not FSODrive.DriveExists("S:") Then
        objNetwork.MapNetworkDrive "S:", "\\SRVHOMOLOGDC1\Medicos","true"
      End If

    Case "DL_Plantonistas"
      If not FSODrive.DriveExists("S:") Then
        objNetwork.MapNetworkDrive "S:", "\\SRVHOMOLOGDC1\Medicos_Plantonistas","true"
      End If

  End Select
Next


'==========================================================================='
' Mapear Impressoras (Mapeamento também pode ser realizado via GPO)         '
'==========================================================================='

set WshNetwork = WScript.CreateObject("WScript.Network")
WshNetwork.AddWindowsPrinterConnection "\\SRVHOMOLOGDC1\Brother", "Brother"
WshNetwork.AddWindowsPrinterConnection "\\SRVHOMOLOGDC1\HP", "HP"
'WshNetwork.SetDefaultPrinter "\\SRVHOMOLOGDC1\Brother", "Brother" '


'==========================================================================='
' Mapear Pastas                                                             '
'==========================================================================='

WshNetwork.MapNetworkDrive "P:", "\\SRVHOMOLOGDC1\Publica", "true"
WshNetwork.MapNetworkDrive "E:", "\\SRVHOMOLOGDC1\Digitalizacoes", "true"


'==========================================================================='
' Criar Atalho para um Site no Desktop                                      '
'==========================================================================='

set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")

set oUrlLink = WshShell.CreateShortcut(strDesktop & "\RD Web Access.lnk")

oUrlLink.TargetPath = "http://app01.system.com.br/RDWeb/Pages/login.aspx"

oUrlLink.IconLocation = "\\SRVHOMOLOGDC1\Icones\favicon.ico"

oUrlLink.Save


'==========================================================================='
' Criar Atalho do Compartilhamento no Desktop                               '
'==========================================================================='

strAppPath = "S:\"
set wshShell = CreateObject("WScript.Shell")
objDesktop = wshShell.SpecialFolders("Desktop")
set oShellLink = wshShell.CreateShortcut(objDesktop & "\Pasta_do_Departamento.lnk")
oShellLink.TargetPath = strAppPath
oShellLink.WindowStyle = "1"
oShellLink.Description = "Pasta_do_Departamento"
oShellLink.Save

strAppPath = "P:\"
set wshShell = CreateObject("WScript.Shell")
objDesktop = wshShell.SpecialFolders("Desktop")
set oShellLink = wshShell.CreateShortcut(objDesktop & "\Pasta_Publica.lnk")
oShellLink.TargetPath = strAppPath
oShellLink.WindowStyle = "1"
oShellLink.Description = "Pasta_Publica"
oShellLink.Save

'Envia o comando para apertar a tecla F5 para atualizar os ícones no Desktop '
WshShell.SendKeys "{F5}"


'==========================================================================='
' Mensagem no logon                                                         '
'==========================================================================='

'set objUser = WScript.CreateObject("WScript.Network") '
'wuser = objUser.UserName '
'  If Time <= "12:00:00" Then '
'    MsgBox ("Bom Dia "+wuser+", você acaba de ingressar na rede corporativa da Hospital X, por favor respeite as políticas de segurança e bom trabalho!") '

'  ElseIf Time >= "12:00:01" And Time <= "18:00:00" Then '
'    MsgBox ("Boa Tarde "+wuser+", você acaba de ingressar na rede corporativa da Hospital X, por favor respeite as políticas de segurança e bom trabalho!") '

'  Else '
'    MsgBox ("Boa Noite "+wuser+", você acaba de ingressar na rede corporativa da Hospital X, por favor respeite as políticas de segurança e bom trabalho!") '

'  End If '


MsgBox ("ATENÇÃO: Pedimos que ao desligar seu computador, escolha a opção Instalar as atualizações e desligar. " & vbcrlf & "Somente assim seu computador instalará atualizações críticas de segurança e ficará atualizado e seguro. " & vbcrlf & "Agradecemos a compreensão, " & vbcrlf & "Equipe da TI")


WScript.Quit