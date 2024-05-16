Attribute VB_Name = "Module1"
Sub Raspredelenie()

Dim coll As New Collection
Dim coll_temp As New Collection
Dim num_of_people As Integer
Dim temp_poz, temp_item As Integer

Range("A2:B1100").Clear
num_of_people = Range("H1").Value

'�������� ����� ����������
If num_of_people > 1 Then


    '���������� ����� � ������� �
    For i = 1 To num_of_people
        Cells(i + 1, 1) = i
    Next
    
    
    '���������� ���������
    For i = 1 To num_of_people
        coll.Add (i)
    Next
    
    '���������� ���������
    For i = 1 To coll.Count
    
        '���������� ��������� ������� �� ���������
        temp_poz = Application.RoundUp(coll.Count * Rnd, 0)
        
        '���������� ������� ���� �������
        temp_item = coll.Item(temp_poz)
        
        '���� ������� ����� ������� �������, �� ���������� ������
        If temp_item = i Then
            Do While temp_item = i
                temp_item = coll.Item(Application.RoundUp(coll.Count * Rnd, 0))
            Loop
        End If
        
        '���� ������� ������� � ���������
        For k = 1 To coll.Count
            '���� �����
            If coll.Item(k) = temp_item Then
                
                '���� �������� 2 ��������
                If coll.Count = 2 Then
                    
                   '���� ��� ��������� ��������� � ����� ������
                   If coll.Item(2) = num_of_people Then
                        
                        '��������� �� ��������� ���������
                        coll_temp.Add (coll.Item(2))
                        Cells(i + 1, 2) = coll.Item(2)
                        '������� ������� �� ������� ���������
                        coll.Remove (2)
                        Exit For
                        
                    Else
                    
                        '��������� �� ��������� ���������
                        coll_temp.Add (coll.Item(k))
                        Cells(i + 1, 2) = temp_item
                        
                        '������� ������� �� ������� ���������
                        coll.Remove (k)
                        Exit For
                    
                    End If
    
                Else
                    
                    '��������� �� ��������� ���������
                    coll_temp.Add (coll.Item(k))
                    Cells(i + 1, 2) = temp_item
                    
                    '������� ������� �� ������� ���������
                    coll.Remove (k)
                    Exit For
                    
                End If
                
            End If
        
        Next
    
        
    Next

Else

    MsgBox ("��������� ������ �������� ���������� ������, ��� 1")

End If

End Sub


