Option Explicit
Dim file_sys_obj , obj, Actual_Time,read_file
Set file_sys_obj = CreateObject("Scripting.FileSystemObject")
Set obj = file_sys_obj.OpenTextFile("C:\Users\pa.puja\Music\OutputTimeStored.txt",8)
Set read_file = file_sys_obj.OpenTextFile("C:\Users\pa.puja\Music\samp - Copy.vbs",1)
Execute read_file.ReadAll()

While flag = False
Actual_Time = Time()
MsgBox Actual_Time
obj.WriteBlankLines(1)
obj.Write Actual_Time 

WScript.Sleep(300000)  '5 minutes delay
Actual_Time = Time()

obj.WriteBlankLines(1)
obj.Write Actual_Time 
Wend
read_file.close

'To append time into TimeStored.txt for every 5 mins until samp - Copy file gets closed..
