@echo off

rem https://duck.co/help/results/syntax


setlocal EnableDelayedExpansion

set S1=%1
if defined S1 (
  set S1=%S1:"=\"%
)

set S2=%2
if defined S2 (
  set S2=%S2:"=\"%
)

set S3=%3
if defined S3 (
  set S3=%S3:"=\"%
)

rem echo Searching Office Dev Guide for %1 %2 %3...
rem echo Searching Office Dev Guide for %S1% %S2% %S3%...


rem lookup the top-level Office Dev Guide if no input args
if [%1]==[] (
  start https://wiki.openoffice.org/wiki/Documentation/DevGuide/OpenOffice.org_Developers_Guide
  exit /B
)



rem Map Writer & text keywords to suitable site restrictions
IF /I [%S1%]==[writer] GOTO WRITER_SEARCH
if /I [%S1%]==[text]  GOTO WRITER_SEARCH
GOTO NO_WRITER

:WRITER_SEARCH
if [%S2%]==[] (
  start https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Text_Documents
) else (
  set SITE=wiki.openoffice.org/wiki/Documentation/DevGuide/Text/
  start http://duckduckgo.com/?q=\\%S2%+%S3%+site:!SITE!
)
exit /B
:NO_WRITER



rem Map Calc & sheet keywords to suitable site restrictions
IF /I [%S1%]==[calc] GOTO CALC_SEARCH
if /I [%S1%]==[spreadsheet] GOTO CALC_SEARCH
if /I [%S1%]==[sheet] GOTO CALC_SEARCH
GOTO NO_CALC

:CALC_SEARCH
if [%S2%]==[] (
  start https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Spreadsheet_Documents
) else (
  set SITE=wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/
  start http://duckduckgo.com/?q=\\%S2%+%S3%+site:!SITE!
)
exit /B
:NO_CALC



rem Map Draw, Impress & keywords to suitable site restrictions
IF /I [%S1%]==[draw] GOTO DRAW_SEARCH
if /I [%S1%]==[impress] GOTO DRAW_SEARCH
if /I [%S1%]==[slide] GOTO DRAW_SEARCH
if /I [%S1%]==[slides] GOTO DRAW_SEARCH
GOTO NO_DRAW

:DRAW_SEARCH
if [%S2%]==[] (
  start https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Drawing_Documents_and_Presentation_Documents
) else (
  set SITE=wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/
  start http://duckduckgo.com/?q=\\%S2%+%S3%+site:!SITE!
)
exit /B
:NO_DRAW



rem Map Chart & keywords to suitable site restrictions
IF /I [%S1%]==[chart] GOTO CHART_SEARCH
IF /I [%S1%]==[charts] GOTO CHART_SEARCH
if /I [%S1%]==[graph] GOTO CHART_SEARCH
GOTO NO_CHART

:CHART_SEARCH
if [%S2%]==[] (
  start https://wiki.openoffice.org/wiki/Documentation/DevGuide/Charts/Charts
) else (
  set SITE=wiki.openoffice.org/wiki/Documentation/DevGuide/Charts/
  start http://duckduckgo.com/?q=\\%S2%+%S3%+site:!SITE!
)
exit /B
:NO_CHART




rem Map Base & keywords to suitable site restrictions
IF /I [%S1%]==[base] GOTO BASE_SEARCH
if /I [%S1%]==[database] GOTO BASE_SEARCH
GOTO NO_BASE

:BASE_SEARCH
echo base
if [%S2%]==[] (
  start https://wiki.openoffice.org/wiki/Documentation/DevGuide/Database/Database_Access
) else (
  echo bar
  set SITE=wiki.openoffice.org/wiki/Documentation/DevGuide/Database/
  start http://duckduckgo.com/?q=\\%S2%+%S3%+site:!SITE!
)
exit /B
:NO_BASE




rem Map Form to suitable site restrictions
IF /I [%S1%]==[form] GOTO FORM_SEARCH
IF /I [%S1%]==[forms] GOTO FORM_SEARCH
GOTO NO_FORM

:FORM_SEARCH
if [%S2%]==[] (
  start https://wiki.openoffice.org/wiki/Documentation/DevGuide/Forms/Forms
) else (
  set SITE=wiki.openoffice.org/wiki/Documentation/DevGuide/Forms/
  start http://duckduckgo.com/?q=\\%S2%+%S3%+site:!SITE!
)
exit /B
:NO_FORM

set SITE=wiki.openoffice.org/wiki/Documentation/DevGuide/
start http://duckduckgo.com/?q=\\%S1%+%S2%+%S3%+site:%SITE%

