Option Explicit
'create some test data
Private Function getTestData() As cJobject
    
    ' just get some vanilla json data
    Set getTestData = JSONParse("[{'name':'john','age':25},{'age':50,'name':'mary'}]")

End Function
Private Function getMoreData() As cJobject
    
    ' just get some vanilla json data
    Set getMoreData = JSONParse("[" & _
            "{'name':'john', 'demographics':{'age':25,'sex':'male'}}," & _
            "{'name':'mary', 'demographics':{'age':50,'sex':'female'}}" & _
        "]")

End Function
' simple test
Private Sub showData()
    Dim job As cJobject, jo As cJobject, jp As cJobject, jm As cJobject
    
    Set job = getMoreData()
    
    ' stringify
    Debug.Print job.stringify
    
    ' iterate
    For Each jo In job.children
        ' jo is each array member
        Debug.Print jo.key, jo.childIndex
        For Each jp In jo.children
            ' jp is each property in each object in the array
            Debug.Print jp.key, jp.value, jo.childIndex, jp.fullKey
        Next jp
    Next jo
    
    ' make an empty one
    Set jm = New cJobject
    With jm.init(Nothing)
        'add an item and make it an array
        With .add.addArray
            ' add an array member
            With .add
                ' add an object to this array member
                .add "name", "john"
                .add "age", 25
            End With
            ' add an array member
            With .add
                ' add an object to this array member
                .add "name", "mary"
                .add "age", 50
            End With
        End With
    End With
    Debug.Print jm.stringify
    ' clean up
    job.tearDown
    jm.tearDown
    
End Sub

' simple test
Private Sub showAccessing()
    Dim job As cJobject, child As cJobject, jp As cJobject, jm As cJobject
    
    Set job = getTestData()
    
    ' the first array member
    Set child = job.children(1)
    Debug.Print child.stringify
    
    Debug.Print job.children(1).child("name").value
    
    ' equivalent to
    Debug.Print job.child("1.name").value
    
    Debug.Print job.child("1.name").stringify
    
End Sub

' simple test
Private Sub showFinding()
    Dim job As cJobject, child As cJobject, jp As cJobject, jm As cJobject
    
    Set job = getMoreData()
    
    ' the first match
    'Set child = job.find("name")
    'Debug.Print child.fullKey
    'Debug.Print child.stringify
    
    'Set child = job.find("2.demographics.age")
    'Debug.Print child.fullKey
    'Debug.Print child.stringify
    
    'Debug.Print job.find("2.demographics.age").parent.child("sex").stringify
    
    'Debug.Print job.find("2.demographics.age").parent.parent.child("name").stringify
    
    'Debug.Print job.findInArray("name", "Mary").parent.stringify

    Debug.Print job.findByValue(50).parent.parent.stringify
    
End Sub
' simple test
Private Sub showObjects()
    Dim job As cJobject, child As cJobject, jp As cJobject, jm As cJobject, headings As cJobject
    
    Set job = getTestData()
    Debug.Print job.stringify
    
    
    Dim r As Range
    Set r = Range("Sheet1!a1")
    
    ' write the headings
    Set headings = New cJobject
    With headings.init(Nothing)
        For Each jm In job.children(1).children
            headings.add jm.key, r.Offset(, jm.childIndex - 1)
            headings.getObject(jm.key).value = jm.key
        Next jm
    End With
    Debug.Print headings.stringify
    
    ' now the data
    For Each jp In job.children
        For Each jm In jp.children
            headings.getObject(jm.key).Offset(jp.childIndex).value = jm.value
        Next jm
    Next jp
    
    For Each jp In headings.children
        Debug.Print jp.key, jp.getObject().Address
    Next jp
    ' the first match
    'Set child = job.find("name")
    'Debug.Print child.fullKey
    'Debug.Print child.stringify
    
    'Set child = job.find("2.demographics.age")
    'Debug.Print child.fullKey
    'Debug.Print child.stringify
    
    'Debug.Print job.find("2.demographics.age").parent.child("sex").stringify
    
    'Debug.Print job.find("2.demographics.age").parent.parent.child("name").stringify
    
    'Debug.Print job.findInArray("name", "Mary").parent.stringify

    'Debug.Print job.findByValue(50).parent.parent.stringify
    
End Sub