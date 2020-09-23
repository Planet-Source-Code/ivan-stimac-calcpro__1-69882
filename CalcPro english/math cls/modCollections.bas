Attribute VB_Name = "modCollections"
Public Function isInCollection(ByVal fName As String, ByRef mCol As Collection) As Boolean
    
    On Error Resume Next
    Dim k As String

    k = mCol.Item(fName)
    
    If Err.Number = 0 Then
        isInCollection = True
        
    Else
        isInCollection = False
        
    End If
    

End Function
