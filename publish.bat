robocopy . "\\IIHS\IIHSDrive\VRC\Private\DIAdem\crashworthiness" /MIR /Z /XD ".git" ".vs" ".vscode" /XF "diadem.sln" "publish.bat" "vwd.webinfo" "web.config" "web.debug.config"
IF %ERRORLEVEL% LEQ 2 exit 0