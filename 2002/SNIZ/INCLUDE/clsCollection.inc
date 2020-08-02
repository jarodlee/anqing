<%
Class clsCollection

 Private m_objItems

 '-------------------------------------------------------------------------------
 '- Event       : Class_Initialize 
 '- Description : Initialization when the object is created
 '-------------------------------------------------------------------------------
 Private Sub Class_Initialize()
   Set m_objItems = Server.CreateObject("Scripting.Dictionary")
   m_objItems.CompareMode = vbTextCompare
 End Sub

 '-------------------------------------------------------------------------------
 '- Property    : Count (Get)
 '- Description : Returns the number of items stored in this object
 '-------------------------------------------------------------------------------
 Public Property Get Count()
   Count = m_objItems.Count
 End Property

 '-------------------------------------------------------------------------------
 '- Property    : Item (Get)
 '- Input       : Index - Textual or integer Id of the requested item
 '- Output      : The requested item or Nothing if tan invalid index was given
 '- Description : Returns a requested item stored in this object
 '-------------------------------------------------------------------------------
 Public Function Item( Index )
 Dim Items
 Dim Value
   Items = m_objItems.Items
   If IsNumeric( Index ) Then
     If ( Index > m_objItems.Count ) Then
       Set Value = Nothing
     Else
       If IsObject(Items(Index)) Then
         Set Value = Items(Index)
       Else
         Value = Items(Index)
       End If
     End If
   Else
     If IsObject(m_objItems(Index)) Then
       Set Value = m_objItems( Index )
     Else
       Value = m_objItems( Index )
     End If
   End If
   If IsObject(Value) Then
     Set Item = Value
   Else
     Item = Value
   End If
 End Function

 '-------------------------------------------------------------------------------
 '- Property    : Key (Get)
 '- Input       : Index - Numerical Id of the item for which the key value is
 '-                       requested
 '- Output      : The requested key value
 '- Description : Gives the key of the Index'th item stored within this object
 '-------------------------------------------------------------------------------
 Public Property Get Key( Index )
   If IsNumeric( Index ) Then
     Key = m_objItems( Index )
   End If
 End Property

 '-------------------------------------------------------------------------------
 '- Function    : Add
 '- Input       : Name  - Key for the value to store
 '-               Value - Value of the item indicated by the given key
 '- Output      : Returns whether the item was successfully added to the list
 '-               or not
 '- Description : Function to add an item to the item-list
 '-------------------------------------------------------------------------------
 Public Function Add(Name, Value)
   If m_objItems.Exists( Name ) Then
     Add = False
   Else
     m_objItems.Add Name, Value
     Add = True
   End If
 End Function

 '-------------------------------------------------------------------------------
 '- Function    : Remove
 '- Input       : Index - Textual or numerical Id of the item to delete
 '- Description : Deletes an item indicated by a given Id
 '-------------------------------------------------------------------------------
 Public Sub Remove( Index )
   If IsNumeric( Index ) Then
     m_objItems.Remove( m_objItems.Item(Index) )
   Else
     m_objItems.Remove( Index )
   End If
 End Sub

 '-------------------------------------------------------------------------------
 '- Function    : RemoveAll
 '- Description : Removes all items from the item-list
 '-------------------------------------------------------------------------------
 Public Sub RemoveAll()
   m_objItems.RemoveAll
 End Sub

End Class
%>