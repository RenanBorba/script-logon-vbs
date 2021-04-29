# Script de Logon VBScript
## Microsoft Windows Server 2016  






### Impedindo a Exibição de Erro para o Usuário                               


On error Resume Next <br>
Err.clear 0 <br>



### Sincroniza o Horário da Estação com o Servidor                                   


'set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") ' <br>
'set objShell = CreateObject("WScript.shell") ' <br>
'strCmd = "net time \\nomeserver /set /yes" ' <br>
'set objExec = objShell.exec(strCmd) ' <br><br>



### Mapear Pastas de acordo com o Grupo do USER                               


set objNetwork = CreateObject("WScript.Network") <br>
strDom = objNetwork.UserDomain <br>
strUser = objNetwork.UserName <br>
set objUser = GetObject("WinNT://" & strDom & "/" & strUser & ",user") <br>
set FSODrive = CreateObject("Scripting.FileSystemObject") <br>

For Each objGroup In objUser.Groups <br>

  Select Case objGroup.Name <br>
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

  End Select <br>
Next <br><br>



## Mapear Impressoras (Mapeamento também pode ser realizado via GPO)         


set WshNetwork = WScript.CreateObject("WScript.Network") <br>
WshNetwork.AddWindowsPrinterConnection "\\SRVHOMOLOGDC1\Brother", "Brother" <br>
WshNetwork.AddWindowsPrinterConnection "\\SRVHOMOLOGDC1\HP", "HP" <br>
'WshNetwork.SetDefaultPrinter "\\SRVHOMOLOGDC1\Brother", "Brother" ' <br><br>



## Mapear Pastas                                                             


WshNetwork.MapNetworkDrive "P:", "\\SRVHOMOLOGDC1\Publica", "true" <br>
WshNetwork.MapNetworkDrive "E:", "\\SRVHOMOLOGDC1\Digitalizacoes", "true" <br>



### Criar Atalho para um Site no Desktop                                      


set WshShell = WScript.CreateObject("WScript.Shell") <br>
strDesktop = WshShell.SpecialFolders("Desktop") <br>

set oUrlLink = WshShell.CreateShortcut(strDesktop & "\RD Web Access.lnk") <br>

oUrlLink.TargetPath = "http://app01.system.com.br/RDWeb/Pages/login.aspx" <br>

oUrlLink.IconLocation = "\\SRVHOMOLOGDC1\Icones\favicon.ico" <br>

oUrlLink.Save <br>



### Criar Atalho do Compartilhamento no Desktop                               


strAppPath = "S:\" <br>
set wshShell = CreateObject("WScript.Shell") <br>
objDesktop = wshShell.SpecialFolders("Desktop") <br>
set oShellLink = wshShell.CreateShortcut(objDesktop & "\Pasta_do_Departamento.lnk") <br>
oShellLink.TargetPath = strAppPath <br>
oShellLink.WindowStyle = "1" <br>
oShellLink.Description = "Pasta_do_Departamento" <br>
oShellLink.Save <br>

strAppPath = "P:\" <br>
set wshShell = CreateObject("WScript.Shell") <br>
objDesktop = wshShell.SpecialFolders("Desktop") <br>
set oShellLink = wshShell.CreateShortcut(objDesktop & "\Pasta_Publica.lnk") <br>
oShellLink.TargetPath = strAppPath <br>
oShellLink.WindowStyle = "1" <br>
oShellLink.Description = "Pasta_Publica" <br>
oShellLink.Save <br>

'Envia o comando para apertar a tecla F5 para atualizar os ícones no Desktop ' <br>
WshShell.SendKeys "{F5}" <br><br>



### Mensagem no logon                                                         


'set objUser = WScript.CreateObject("WScript.Network") ' <br>
'wuser = objUser.UserName ' <br>
'  If Time <= "12:00:00" Then ' <br>
'    MsgBox ("Bom Dia "+wuser+", você acaba de ingressar na rede corporativa da Hospital X, por favor respeite as políticas de segurança e bom trabalho!") ' <br><br>

'  ElseIf Time >= "12:00:01" And Time <= "18:00:00" Then ' <br>
'    MsgBox ("Boa Tarde "+wuser+", você acaba de ingressar na rede corporativa da Hospital X, por favor respeite as políticas de segurança e bom trabalho!") ' <br>

'  Else ' <br>
'    MsgBox ("Boa Noite "+wuser+", você acaba de ingressar na rede corporativa da Hospital X, por favor respeite as políticas de segurança e bom trabalho!") ' <br>

'  End If ' <br>


MsgBox ("ATENÇÃO: Pedimos que ao desligar seu computador, escolha a opção Instalar as atualizações e desligar. " & vbcrlf & "Somente assim seu computador instalará atualizações críticas de segurança e ficará atualizado e seguro. " & vbcrlf & "Agradecemos a compreensão, " & vbcrlf & "Equipe da TI") <br><br>


<br>
WScript.Quit
