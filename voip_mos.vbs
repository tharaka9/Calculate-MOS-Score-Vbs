Set args = Wscript.Arguments

dateTime = Now()
' Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("D:\New folder\vbs\voipLog.log",8,true) 'add path to save log file

For Each arg In args
  ip = arg
Next

' objFileToWrite.WriteLine(dateTime&"  Ip: " &ip)

Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec("ping -nc 5 " & ip)
strPingResult = objExec.StdOut.ReadAll
arrayPingResult = Split(strPingResult, " ")

' objFileToWrite.WriteLine(dateTime&"  Ping result: " &strPingResult)

packet_loss = arrayPingResult(47)
avarage = arrayPingResult(66)

' get time
test =  arrayPingResult(10)
test1 =  arrayPingResult(15)
test2 =  arrayPingResult(20)
test3 =  arrayPingResult(25)
test4 =  arrayPingResult(30)

Set re = New RegExp 'remove letters from data
re.Pattern = "[^0-9\.,]"
re.Global = True
re.IgnoreCase = True

data1 = cint(re.Replace(test, ""))
data2 = cint(re.Replace(test1, ""))
data3 = cint(re.Replace(test2, ""))
data4 = cint(re.Replace(test3, ""))
data5 = cint(re.Replace(test4, ""))
avg_latency = cint(re.Replace(avarage, ""))

' objFileToWrite.WriteLine(dateTime&"  Time 1: " &data1)
' objFileToWrite.WriteLine(dateTime&"  Time 2: " &data2)
' objFileToWrite.WriteLine(dateTime&"  Time 3: " &data3)
' objFileToWrite.WriteLine(dateTime&"  Time 4: " &data4)
' objFileToWrite.WriteLine(dateTime&"  avg_latency: " &avarage)



sub1 = data1 - data2
sub2 = data2 - data3
sub3 = data3 - data4
sub4 = data4 - data5


' objFileToWrite.WriteLine(dateTime&"  TimeDef 1: " &sub1)
' objFileToWrite.WriteLine(dateTime&"  TimeDef 2: " &sub2)
' objFileToWrite.WriteLine(dateTime&"  TimeDef 3: " &sub3)
' objFileToWrite.WriteLine(dateTime&"  TimeDef 4: " &sub4)



sub1 = Abs(sub1) 'convert negative value to positive
sub2 = Abs(sub2)
sub3 = Abs(sub3)
sub4 = Abs(sub4)
tot = sub1+sub2+sub3+sub4

' objFileToWrite.WriteLine(dateTime&"  Total: " &tot)

jitter = tot / 4

' objFileToWrite.WriteLine(dateTime&"  Jitter: " &jitter&vbCrLf)
' objFileToWrite.WriteLine("......................................................")

effective_latency = (avg_latency + jitter * 2 + 10)
If effective_latency < 160 Then
r_value = 93.2 - (effective_latency / 40)
Else
r_value = 93.2 - (effective_latency - 120) / 10
End If
r_value = r_value - (packet_loss * 2.5)


If r_value < 0 Then
r_value = 0
ret_val = 1 + (0.035) * r_value + (0.000007) * (r_value) * (r_value - 60) * (100 - r_value)
ElseIf r_value > 100 Then
r_value = 100
ret_val = 1 + (0.035) * r_value + (0.000007) * (r_value) * (r_value - 60) * (100 - r_value)
Else
ret_val = 1 + (0.035) * r_value + (0.000007) * (r_value) * (r_value - 60) * (100 - r_value)
End If



ret_val = round(ret_val, 1)   
avg_latency = round(avg_latency, 0)
jitter = round(jitter, 1)

' wscript.echo "Mos: " &ret_val
' wscript.echo "Avarage Latency: " &avg_latency
' wscript.echo "Packet Loss: " &packet_loss
' wscript.echo "Jitter: " &jitter 

' Avg latency = 873, jitter = 1,523.8 MOS = 262.8 which is incorrect value.



' objFileToWrite.Close

Wscript.StdOut.WriteLine "<module>"
Wscript.StdOut.WriteLine "    <name><![CDATA["& ip &"_Voip_Mos]]></name>"
Wscript.StdOut.WriteLine "    <type><![CDATA[generic_data]]></type>"
Wscript.StdOut.WriteLine "    <description><![CDATA[Mean Opinion Score, is a measure of voice quality]]></description>"
Wscript.StdOut.WriteLine "    <min_critical><![CDATA[0]]></min_critical>"
Wscript.StdOut.WriteLine "    <max_critical><![CDATA[3.1]]></max_critical>"
Wscript.StdOut.WriteLine "    <data><![CDATA[" & ret_val & "]]></data>"
Wscript.StdOut.WriteLine "</module>"

Wscript.StdOut.WriteLine "<module>"
Wscript.StdOut.WriteLine "    <name><![CDATA["& ip &"_Voip_Latency]]></name>"
Wscript.StdOut.WriteLine "    <type><![CDATA[generic_data]]></type>"
Wscript.StdOut.WriteLine "    <description><![CDATA[Mean Opinion Score, is a measure of voice quality Latency]]></description>"
Wscript.StdOut.WriteLine "    <data><![CDATA[" & avg_latency & "]]></data>"
Wscript.StdOut.WriteLine "</module>"

Wscript.StdOut.WriteLine "<module>"
Wscript.StdOut.WriteLine "    <name><![CDATA["& ip &"_Voip_Jitter]]></name>"
Wscript.StdOut.WriteLine "    <type><![CDATA[generic_data]]></type>"
Wscript.StdOut.WriteLine "    <description><![CDATA[Mean Opinion Score, is a measure of voice quality Jitter]]></description>"
Wscript.StdOut.WriteLine "    <min_critical><![CDATA[0]]></min_critical>"
Wscript.StdOut.WriteLine "    <max_critical><![CDATA[30]]></max_critical>"
Wscript.StdOut.WriteLine "    <data><![CDATA[" & jitter & "]]></data>"
Wscript.StdOut.WriteLine "</module>"

Wscript.StdOut.WriteLine "<module>"
Wscript.StdOut.WriteLine "    <name><![CDATA["& ip &"_Voip_Packet_Loss]]></name>"
Wscript.StdOut.WriteLine "    <type><![CDATA[generic_data]]></type>"
Wscript.StdOut.WriteLine "    <description><![CDATA[Mean Opinion Score, is a measure of voice quality Packet Loss]]></description>"
Wscript.StdOut.WriteLine "    <min_critical><![CDATA[0]]></min_critical>"
Wscript.StdOut.WriteLine "    <max_critical><![CDATA[1]]></max_critical>"
Wscript.StdOut.WriteLine "    <data><![CDATA[" & packet_loss & "]]></data>"
Wscript.StdOut.WriteLine "</module>"
