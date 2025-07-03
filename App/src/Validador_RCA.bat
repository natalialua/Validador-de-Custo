@echo off
setlocal

:: Caminho do executável na rede
set FONTE="G:"

:: Caminho destino: Área de Trabalho do usuário
set FONTE="G:"

:: Copia o .exe da rede para a Área de Trabalho (somente se for mais recente)
xcopy %FONTE% %DESTINO% /D /Y /Q

:: Executa o .exe a partir da Área de Trabalho
start "" %DESTINO%

endlocal
exit
