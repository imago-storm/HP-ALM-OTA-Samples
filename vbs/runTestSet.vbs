

' To run: C:\Windows\SysWOW64\cscript.exe  .\runTestSet.vbs

Set QCConnection = CreateObject("TDApiOle80.TDConnection")

Dim sUserName, sPassword
sUserName = "admin"
sPassword = "changeme"
domain = "DEFAULT"
project = "TestProject"
nPath = "Root\QA"
tSetName = "SampleTestSet"

QCConnection.InitConnectionEx "http://localhost:8080/qcbin"
QCConnection.Login sUserName, sPassword
QCConnection.Connect domain, project

WScript.Echo "Connected"


Set TSetFact = QCConnection.TestSetFactory
Set tsTreeMgr = QCConnection.TestSetTreeManager


Set tsFolder = tsTreeMgr.NodeByPath(nPath)

If tsFolder Is Nothing Then
    err.Raise vbObjectError + 1, "RunTestSet", "Could not find folder " & nPath
End If

Set tsList = tsFolder.FindTestSets(tSetName)
If tsList.Count > 1 Then
    MsgBox "FindTestSets found more than one test set: refine search"
ElseIf tsList.Count < 1 Then
    MsgBox "FindTestSets: test set not found"
End If


Set theTestSet = tsList.Item(1)
Wscript.Echo theTestSet.ID


Set Scheduler = theTestSet.StartExecution("")

Scheduler.RunAllLocally = True

Scheduler.Run
' Get the execution status object.
Set execStatus = Scheduler.ExecutionStatus

Dim runFinished, iter

While ((RunFinished = False) And (iter < 100))
    iter = iter + 1
    execStatus.RefreshExecStatusInfo "all", True
    RunFinished = execStatus.Finished
    Set EventsList = execStatus.EventsList

    For i = 1 To execStatus.Count
        Set TestExecStatusObj = execStatus.Item(i)

        WSCript.Echo "Iteration " & iter & " Status: " & TestExecStatusObj.Message & " " & TestExecStatusObj.Status


    Next
    'Sleep() has to be declared before it can be used.
    'This is the module level declaration of Sleep():
    'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    WSCript.Sleep 5000

Wend 'Loop While execStatus.Finished = False


WScript.Echo "Finished"




