Attribute VB_Name = "ConvertNumber"

' Macro สำหรับแปลงเลข อาราบิก เป็น เลขไทย
Sub Arabic2thai()
    Dim i As Integer
    For i = 0 To 9
        With Selection.Find
            .Text = Chr(48 + i)
            .Replacement.Text = ChrW(3664 + i)
            .Wrap = wdFindContinue
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub

' Macro สำหรับแปลงเลขไทย เป็น เลขอาราบิก
Sub Thai2Arabic()
    Dim i As Integer
    For i = 0 To 9
        With Selection.Find
            .Text = ChrW(3664 + i)
            .Replacement.Text = Chr(48 + i)
            .Wrap = wdFindContinue
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub
