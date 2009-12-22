<%
'
' ---------------------------------------------------------------------------
'    $Source: /home/cvs/MeCMS/src/MeList.asp,v $
'    $Revision: 1.5 $
'    $Author: riceball $
' ---------------------------------------------------------------------------
'
' Implements a resizable List class.
'
' Properties:
'   Items(pIndex: Integer)
'   Capacity: Integer
'   Delimiter: the default is comma(",")
'   DelimitedText
'   readonly Count: Integer
'   readonly List: TArray: return the FList!
' Methods:
'   Sub Pack: reduce the Capacity to the Count
'   Function Add(aValue):Integer  add aValue to the list, return the index in list, or return -1.


'the Duplicates constants
Const dupIgnore = 0
Const dupAccept = 1
Const dupError  = 2

Class TMeList
    Private FList
    Private FCount

    Public  Delimiter
    Public Duplicates

    Private Sub Class_Initialize()
        Redim FList(8)
        FCount = 0
        Delimiter = ","
        Duplicates = dupIgnore
    End Sub

    Private Sub Class_Terminate()
    End Sub

    Public Property Let DelimitedText(pValue)
        Dim vTempList, vItem
        if Duplicates = dupAccept then
          FList = Split(pValue, Delimiter, -1, vbTextCompare)
          FCount = UBound(FList) + 1
        else
          vTempList = Split(pValue, Delimiter, -1, vbTextCompare)
          For Each vItem In vTempList
            Add(vItem)
          Next
        end if
    End Property

    Public Property Get DelimitedText()
        Dim i, Result
        if FCount > 0 then
          Result = FList(0)
          For i = 1 to FCount - 1
            Result = Result + Delimiter & FList(i) 
          Next
        else
          Result = ""
        end if
        DelimitedText = Result
    End Property

    'Specifies the allocated size of the array of Values maintained by the TMeList object.
    Public Property Let Capacity(pValue)
        Redim Preserve FList(pValue)
    End Property

    Public Property Get Capacity()
        Capacity = UBound(FList)
    End Property

    Public Property Let Count(pValue)
        if pValue > Capacity then Capacity = pValue
        FCount = pValue
    End Property

    Public Property Get Count()
        Count = FCount
    End Property

    Public Property Get List()
        List = FList
    End Property

    Public Property Let List(pValue)
        FList = pValue
        'FCount = Capacity
    End Property

    Public Default Property Get Items(pIndex)
        If IsObject(FList(pIndex)) Then
            Set Items = FList(pIndex)
        Else
            Items = FList(pIndex)
        End If
    End Property

    ' set the object value
    Public Property Set Items(pIndex, pObjValue)
        Call SetItem(pIndex, pObjValue)
    End Property

    ' set the common value(not object)
    Public Property Let Items(pIndex, pValue)
        Call SetItem(pIndex, pValue)
    End Property

    Public Sub Pack()
        Dim vC
        vC = FCount - 1
        if vC < 0 then vC = 0
        Redim Preserve FList(vC)
    End Sub

    ' remove all the item = pValue
    Public Sub Trim(pValue)
      Dim Result

        For i = FCount - 1 to 0 Step -1
          if IsObject(pValue) then
            if FList(i) is pValue then Delete(i)
          else
            if FList(i) = pValue then Delete(i)
          end if
        Next
    End Sub

    Public Function Clone()
        Dim Result
        Set Result = New TMeList
        Result.List = FList
        Result.Count = FCount
        Result.Delimiter = Delimiter
        Result.Duplicates = Duplicates
        Set Clone = Result
    End Function

    Public Sub Assign(ByRef aList)
      Clear()
      FList = aList.List
      Count = aList.Count
      Delimiter = aList.Delimiter
      Duplicates = aList.Duplicates
    End Sub

    Public Function Add(pValue)
      Dim Result

        Add = -1
        if Duplicates <> dupAccept then
          Result = IndexOf(pValue)
          if Result >= 0 then 
            if Duplicates = dupError then RaiseError vbListDuplicateError, "TMeList.Add", "Error: this is duplicate in the item of list"
            Exit Function
          end if
        end if
        Result = FCount
        if Result >= Capacity then Grow
        Call SetItem(FCount, pValue)
        FCount = FCount + 1
        Add = Result
    End Function

    Public Sub SetItem(pIndex, pValue)
        If IsObject(pValue) Then
            Set FList(pIndex) = pValue
        Else
            FList(pIndex) = pValue
        End If
    End Sub

    Public Sub Delete(ByVal aIndex)
        FCount = FCount - 1
        Do While aIndex < FCount
            Call SetItem(aIndex, FList(aIndex + 1))
            aIndex = aIndex + 1
        Loop
    End Sub

    Public Sub Insert(pIndex, pValue)
        Dim i
        if FCount <= Capacity then Grow
        i = FCount
        Do While i > pIndex
            Call SetItem(i, FList(i-1))
            i = i - 1
        Loop
        Call SetItem(pIndex, pValue)
        FCount = FCount + 1
    End Sub

    Public Sub Move(pCurIndex, pNewIndex)
      Dim vItem

      if pCurIndex <> pNewIndex then
      begin
        if IsObject(FList(pCurIndex)) then Set vItem = FList(pCurIndex) else vItem = FList(pCurIndex)
        Call Delete(CurIndex)
        Call Insert(pNewIndex, 0)
        Call SetItem(pNewIndex, vItem)
      end if
    End Sub

    'Deletes the first reference to the Item parameter from the Items array.
    'Call Remove to remove a specific item from the Items array when its index is unknown. 
    'The value returned is the index of the item in the Items array before it was removed. 
    'After an item is removed, all the items that follow it are moved up in index position and the Count is reduced by one.

    'If the Items array contains more than one copy of the value, only the first copy is deleted.
    Public Function Remove(ByRef aValue)
      Dim Result
        Result = IndexOf(aValue)
        if Result <> -1 then Call Delete(Result)
        Remove = Result
    End Function

    Public Function IsEmpty()
        IsEmpty = (FCount <= 0)
    End Function

    Public Function IndexOf(pItem)
      Dim i, Result, vItem
        IndexOf = -1
        i = 0
        Do While i < FCount
          if IsObject(FList(i)) then
            Result = (FList(i) is pItem)
          else
            vItem = FList(i)
            if IsArray(vItem) then
              Result = False
              On Error Resume Next
              'writeln(vItem(0)&" "&pItem(0))
              Result = (Join(vItem) = Join(pItem))
              On Error Goto 0
              'Result = IsArray(pItem)
              'On Error Resum Next
              'if Result then Result = (Join(vItem) = Join(pItem))
            else
              Result = (FList(i) = pItem)
            end if
          end if
          if Result then 
            IndexOf = i
            Exit Do
          end if
          i = i + 1
        Loop
    End Function

    Public Function First()
        If IsObject(FList(0)) Then
            Set First = FList(0)
        Else
            First = FList(0)
        End If
    End Function

    Public Function Last()
        Dim vLastIndex
        vLastIndex = FCount-1
        If IsObject(FList(vLastIndex)) Then
            Set Last = FList(vLastIndex)
        Else
            Last = FList(vLastIndex)
        End If
    End Function

    Public Sub Push(pValue)
      Add(pValue)
    End Sub

    Public Function Pop()
        FCount = FCount - 1
        If IsObject(FList(FCount)) Then
            Set Pop = FList(FCount)
        Else
            Pop = FList(FCount)
        End If
    End Function

    Public Function Top()
        If IsObject(Last) Then
            Set Top = Last
        Else
            Top = Last
        End If
    End Function

    Public Sub Clear()
      FCount = 0
    End Sub

    Private Sub Grow()
        Dim Delta

        if Capacity > 63 then
          Delta = Capacity \ 4 - 1
        else
          if Capacity > 7 then
            Delta = 15
          else
            Delta = 3
          end if
        end if
        Capacity = Capacity + Delta
    End Sub
End Class
%>
