Attribute VB_Name = "Module11"
Public Sub ConvertQuestions()
Dim wbSource As Workbook, wbTarget As Workbook
Dim wbAnswers As Workbook, wbOriginal As Workbook
Dim newBook As Workbook
Dim ws As Worksheet
Dim sPath As String, tPath As String
Dim aPath As String, oPath As String
Dim chapter As String, newChapter As String, aChapter As String
Dim i As Integer, k As Integer
Dim qNum As Integer, aNum As String
Dim a As String, b As String
Dim c As String, d As String

tPath = "C:....xls"
oPath = "C:\....xls"
sPath = "C:\.....xlsx"
aPath = "C:\....xlsx"
i = 2
k = 2
qNum = 1
aNum = "001a"
a = "001a"
b = "001b"
c = "001c"
d = "001d"

Set newBook = Workbooks.Add
 With newBook
 .Title = "QuizQuestions"
 .SaveAs Filename:=newBook.Path & tPath
 End With

' Open all four files
Set wbOriginal = Workbooks.Open(oPath)
Set wbSource = Workbooks.Open(sPath)
Set wbTarget = Workbooks.Open(tPath)
Set wbAnswers = Workbooks.Open(aPath)

' Create the first Quiz question and answer
chapter = wbSource.Worksheets("questions").Range("B" & i)
aChapter = wbAnswers.Worksheets("answers").Range("B" & k)
Set ws = wbTarget.Worksheets.Add
ws.Name = chapter
wbOriginal.Worksheets("Sample").Range("A1:R1").Copy _
    Destination:=wbTarget.Worksheets(chapter).Range("A1:R1")
wbSource.Worksheets("questions").Range("D2").Copy _
    Destination:=wbTarget.Worksheets(chapter).Range("B2")
wbAnswers.Worksheets("answers").Range("E2").Copy _
    Destination:=wbTarget.Worksheets(chapter).Range("F2")
k = k + 1
aNum = wbAnswers.Worksheets("answers").Range("D" & k)
Do While chapter = aChapter And qNum = 1
    If aNum = b Then
        wbAnswers.Worksheets("answers").Range("E" & k).Copy _
            Destination:=wbTarget.Worksheets(chapter).Range("G2")
            k = k + 1
            aNum = wbAnswers.Worksheets("answers").Range("D" & k)
            aChapter = wbAnswers.Worksheets("answers").Range("B" & k)
            qNum = wbAnswers.Worksheets("answers").Range("C" & k)
    ElseIf aNum = c Then
        wbAnswers.Worksheets("answers").Range("E" & k).Copy _
            Destination:=wbTarget.Worksheets(chapter).Range("H2")
            k = k + 1
            aNum = wbAnswers.Worksheets("answers").Range("D" & k)
            aChapter = wbAnswers.Worksheets("answers").Range("B" & k)
            qNum = wbAnswers.Worksheets("answers").Range("C" & k)
    ElseIf aNum = d Then
        wbAnswers.Worksheets("answers").Range("E" & k).Copy _
            Destination:=wbTarget.Worksheets(chapter).Range("I2")
            k = k + 1
            aNum = wbAnswers.Worksheets("answers").Range("D" & k)
            aChapter = wbAnswers.Worksheets("answers").Range("B" & k)
            qNum = wbAnswers.Worksheets("answers").Range("C" & k)
    End If
Loop
newChapter = wbSource.Worksheets("questions").Range("B" & i + 1)
' ....

