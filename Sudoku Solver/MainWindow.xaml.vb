Class MainWindow
    Public Const GridSize As Integer = 9
    Public Property TextBoxes As New List(Of TextBox)
    Private Sub Run_Button_Click(sender As Object, e As RoutedEventArgs)
        'GridArray(Row, Column)
        Dim GridArray(8, 8) As Integer
        GridArray = GetGridArrayData()

        If GetAllValues(GridArray) Then
            SetTextBoxesData(GridArray)
        Else
            MessageBox.Show("Unsolvable")
        End If

    End Sub

    Private Function IsValueInRow(Grid(,) As Integer, Row As Integer, Value As Integer) As Boolean
        For ColCounter As Integer = 0 To 8
            If Grid(Row, ColCounter) = Value Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function IsValueInCol(Grid(,) As Integer, Col As Integer, Value As Integer) As Boolean
        For RowCounter As Integer = 0 To 8
            If Grid(RowCounter, Col) = Value Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function IsValueInSection(Grid(,) As Integer, Row As Integer, Col As Integer, Value As Integer)
        Dim SectionTopRow As Integer = Row - (Row Mod 3)
        Dim SectionLeftCol As Integer = Col - (Col Mod 3)

        For RowCounter As Integer = SectionTopRow To SectionTopRow + 2
            For ColCounter As Integer = SectionLeftCol To SectionLeftCol + 2
                If Grid(RowCounter, ColCounter) = Value Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function

    Private Function IsValueValid(Grid(,) As Integer, Row As Integer, Col As Integer, Value As Integer)
        Return Not (IsValueInRow(Grid, Row, Value)) And
            Not (IsValueInCol(Grid, Col, Value)) And
            Not (IsValueInSection(Grid, Row, Col, Value))
    End Function

    Private Function GetAllValues(Grid(,) As Integer) As Boolean
        For RowCounter As Integer = 0 To 8
            For ColCounter As Integer = 0 To 8
                If Grid(RowCounter, ColCounter) = 0 Then
                    For ValueAttempt As Integer = 1 To 9
                        If IsValueValid(Grid, RowCounter, ColCounter, ValueAttempt) Then
                            Grid(RowCounter, ColCounter) = ValueAttempt
                            If GetAllValues(Grid) = True Then
                                Return True
                            Else
                                Grid(RowCounter, ColCounter) = 0
                            End If
                        End If
                    Next
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    Private Function GetGridArrayData() As Array
        Dim TempGridArray(8, 8) As Integer

        For Row As Integer = 0 To 8
            For Col As Integer = 0 To 8
                Integer.TryParse(TextBoxes((Row * 9) + Col).Text, TempGridArray(Row, Col))
            Next
        Next
        Return TempGridArray
    End Function

    Private Sub SetTextBoxesData(Grid(,) As Integer)
        For Row As Integer = 0 To 8
            For Col As Integer = 0 To 8
                TextBoxes((Row * 9) + Col).Text = Grid(Row, Col)
            Next
        Next
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        'Get list of textboxes on gameboard
        For TextBoxNumber As Integer = 1 To 81
            TextBoxes.Add(Me.FindName($"TextBox_{TextBoxNumber}"))
        Next

    End Sub
End Class
