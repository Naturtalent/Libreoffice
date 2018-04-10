@echo off

rem Find LibreOffice and check that its bit size matches Java's
rem The path is saved to lofindtemp.txt

setlocal

rem get bitness of OS
IF "%ProgramFiles(x86)%"=="" (
  SET OS64=0
  rem  echo OS is 32-bit
) else (
  SET OS64=1
  rem echo OS is 64-bit
)

rem get bitness of Java
java -version 2>&1 | find "64-Bit" >nul:
if %errorlevel%==0 (
  SET Java64=1
  rem echo Java is 64-bit
) else (
  SET Java64=0
  rem echo Java is 32-bit
)


rem find LO and record its bitness
SET LO64=1
rem start by assuming LO is 32 bit

SET LO=%ProgramFiles%\LibreOffice 5
IF EXIST "%LO%" (
  if %OS64%==0 (
    SET LO64=0
  )
  GOTO checkbits
)

SET LO=%ProgramFiles%\LibreOffice 4
IF EXIST "%LO%" (
  if %OS64%==0 (
    SET LO64=0
  )
  GOTO checkbits
)

IF %OS64%==0 (
  rem OS is 32 bit
  echo LibreOffice directory not found; giving up
  EXIT /B
)

rem OS is 64 bit, so look for 32-bit installations of LO
SET LO=%ProgramFiles(x86)%\LibreOffice 5
IF EXIST "%LO%" (
  SET LO64=0
  GOTO checkbits
)

SET LO=%ProgramFiles(x86)%\LibreOffice 4
IF EXIST "%LO%" (
  SET LO64=0
  GOTO checkbits
)


echo LibreOffice directory not found; giving up
EXIT /B


:checkbits


rem Java and LibreOffice should be the same bit sizes
IF %Java64%==1 (
  if %LO64%==1 (
    goto start
  )
)
IF %Java64%==0 (
  if %LO64%==0 (
    goto start
  )
)

echo WARNING: Java and LibreOffice are different bit sizes
if %Java64%==1 (
  echo Java is 64-bit
) else (
  echo Java is 32-bit
)
if %LO64%==1 (
  echo LibreOffice is 64-bit
) else (
  echo LibreOffice is 32-bit
)
echo.

:start

echo Found "%LO%"
echo "%LO%"> lofindTemp.txt
