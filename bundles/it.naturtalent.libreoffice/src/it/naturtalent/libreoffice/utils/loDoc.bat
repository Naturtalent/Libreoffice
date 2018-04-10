@echo off

setlocal

set S1=%1
set S2=%2
set S3=%3

rem lookup the top-level LO module
if "%1"=="" (
  set S1="star"
  set S2="module"
)


rem Map LO appl names to suitable modules
if /I "%1"=="writer" (
  if "%2"=="" (
    set S1=text
    set S2=module
  )
)

if /I "%1"=="draw" (
  if "%2"=="" (
    set S1=drawing
    set S2=module
  )
)

if /I "%1"=="impress" (
  if "%2"=="" (
    set S1=presentation
    set S2=module
  )
)

if /I "%1"=="calc" (
  if "%2"=="" (
    set S1=sheet
    set S2=module
  )
)

if /I "%1"=="chart2" (
  if "%2"=="" (
    set S1=chart2
    set S2=module
  )
)

if /I "%1"=="base" (
  if "%2"=="" (
    set S1=sdbc
    set S2=module
  )
)

rem Suggest a tighter search for non-interface names
set FIRST_CH=%S1:~0,1%
rem echo %FIRST_CH%
if /I not "%FIRST_CH%"=="x" (
 rem echo This is "%S2%"
 if "%S2%"=="" (
    echo Consider adding 'service' or 'module' to the search args
 )
)


echo Searching LO docs for %S1% %S2% %S3%...

rem https://duck.co/help/results/syntax


start http://duckduckgo.com/?q=\%S1%+%S2%+%S3%+site:api.libreoffice.org/docs/idl/ref

:: start http://duckduckgo.com/?q=\%S1%+%S2%+%S3%+site:api.libreoffice.org/docs/idl/ref+intitle:Reference

rem start http://duckduckgo.com/?q=\%S1%+%S2%+%S3%+site:api.libreoffice.org/docs+intitle:-idl

