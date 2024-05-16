Attribute VB_Name = "Module1"
Sub Raspredelenie()

Dim coll As New Collection
Dim coll_temp As New Collection
Dim num_of_people As Integer
Dim temp_poz, temp_item As Integer

Range("A2:B1100").Clear
num_of_people = Range("H1").Value

'проверка числа участников
If num_of_people > 1 Then


    'заполнение ячеек в столбце А
    For i = 1 To num_of_people
        Cells(i + 1, 1) = i
    Next
    
    
    'инициируем коллекцию
    For i = 1 To num_of_people
        coll.Add (i)
    Next
    
    'перебираем коллекцию
    For i = 1 To coll.Count
    
        'определяем случайную позицию из коллекции
        temp_poz = Application.RoundUp(coll.Count * Rnd, 0)
        
        'определяем элемент этой позиции
        temp_item = coll.Item(temp_poz)
        
        'если элемент равен текущей позиции, то перебираем дальше
        If temp_item = i Then
            Do While temp_item = i
                temp_item = coll.Item(Application.RoundUp(coll.Count * Rnd, 0))
            Loop
        End If
        
        'ишем текущий элемент в коллекции
        For k = 1 To coll.Count
            'если нашли
            If coll.Item(k) = temp_item Then
                
                'если остаются 2 элемента
                If coll.Count = 2 Then
                    
                   'если еще последнее совпадает с общим числом
                   If coll.Item(2) = num_of_people Then
                        
                        'добавляем во временную коллекцию
                        coll_temp.Add (coll.Item(2))
                        Cells(i + 1, 2) = coll.Item(2)
                        'удаляем элемент из текущей коллекции
                        coll.Remove (2)
                        Exit For
                        
                    Else
                    
                        'добавляем во временную коллекцию
                        coll_temp.Add (coll.Item(k))
                        Cells(i + 1, 2) = temp_item
                        
                        'удаляем элемент из текущей коллекции
                        coll.Remove (k)
                        Exit For
                    
                    End If
    
                Else
                    
                    'добавляем во временную коллекцию
                    coll_temp.Add (coll.Item(k))
                    Cells(i + 1, 2) = temp_item
                    
                    'удаляем элемент из текущей коллекции
                    coll.Remove (k)
                    Exit For
                    
                End If
                
            End If
        
        Next
    
        
    Next

Else

    MsgBox ("Òðåáóåòñÿ ââåñòè çíà÷åíèå ó÷àñòêèêîâ áîëüøå, ÷åì 1")

End If

End Sub


