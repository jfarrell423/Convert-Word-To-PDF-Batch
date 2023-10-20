echo off
for %%X in (*.docx) do cscript.exe //nologo Convert-word-to-pdf.js "%%X"
for %%X in (*.doc)  do cscript.exe //nologo Convert-word-to-pdf.js "%%X"