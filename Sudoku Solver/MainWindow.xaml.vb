Class MainWindow
    Public Const GridSize As Integer = 9

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

        'I'm know there is a much more elegant way to handel this but it will do for now
        Integer.TryParse(Me.TextBox_1.Text, TempGridArray(0, 0))
        Integer.TryParse(Me.TextBox_2.Text, TempGridArray(0, 1))
        Integer.TryParse(Me.TextBox_3.Text, TempGridArray(0, 2))
        Integer.TryParse(Me.TextBox_4.Text, TempGridArray(0, 3))
        Integer.TryParse(Me.TextBox_5.Text, TempGridArray(0, 4))
        Integer.TryParse(Me.TextBox_6.Text, TempGridArray(0, 5))
        Integer.TryParse(Me.TextBox_7.Text, TempGridArray(0, 6))
        Integer.TryParse(Me.TextBox_8.Text, TempGridArray(0, 7))
        Integer.TryParse(Me.TextBox_9.Text, TempGridArray(0, 8))
        Integer.TryParse(Me.TextBox_10.Text, TempGridArray(1, 0))
        Integer.TryParse(Me.TextBox_11.Text, TempGridArray(1, 1))
        Integer.TryParse(Me.TextBox_12.Text, TempGridArray(1, 2))
        Integer.TryParse(Me.TextBox_13.Text, TempGridArray(1, 3))
        Integer.TryParse(Me.TextBox_14.Text, TempGridArray(1, 4))
        Integer.TryParse(Me.TextBox_15.Text, TempGridArray(1, 5))
        Integer.TryParse(Me.TextBox_16.Text, TempGridArray(1, 6))
        Integer.TryParse(Me.TextBox_17.Text, TempGridArray(1, 7))
        Integer.TryParse(Me.TextBox_18.Text, TempGridArray(1, 8))
        Integer.TryParse(Me.TextBox_19.Text, TempGridArray(2, 0))
        Integer.TryParse(Me.TextBox_20.Text, TempGridArray(2, 1))
        Integer.TryParse(Me.TextBox_21.Text, TempGridArray(2, 2))
        Integer.TryParse(Me.TextBox_22.Text, TempGridArray(2, 3))
        Integer.TryParse(Me.TextBox_23.Text, TempGridArray(2, 4))
        Integer.TryParse(Me.TextBox_24.Text, TempGridArray(2, 5))
        Integer.TryParse(Me.TextBox_25.Text, TempGridArray(2, 6))
        Integer.TryParse(Me.TextBox_26.Text, TempGridArray(2, 7))
        Integer.TryParse(Me.TextBox_27.Text, TempGridArray(2, 8))
        Integer.TryParse(Me.TextBox_28.Text, TempGridArray(3, 0))
        Integer.TryParse(Me.TextBox_29.Text, TempGridArray(3, 1))
        Integer.TryParse(Me.TextBox_30.Text, TempGridArray(3, 2))
        Integer.TryParse(Me.TextBox_31.Text, TempGridArray(3, 3))
        Integer.TryParse(Me.TextBox_32.Text, TempGridArray(3, 4))
        Integer.TryParse(Me.TextBox_33.Text, TempGridArray(3, 5))
        Integer.TryParse(Me.TextBox_34.Text, TempGridArray(3, 6))
        Integer.TryParse(Me.TextBox_35.Text, TempGridArray(3, 7))
        Integer.TryParse(Me.TextBox_36.Text, TempGridArray(3, 8))
        Integer.TryParse(Me.TextBox_37.Text, TempGridArray(4, 0))
        Integer.TryParse(Me.TextBox_38.Text, TempGridArray(4, 1))
        Integer.TryParse(Me.TextBox_39.Text, TempGridArray(4, 2))
        Integer.TryParse(Me.TextBox_40.Text, TempGridArray(4, 3))
        Integer.TryParse(Me.TextBox_41.Text, TempGridArray(4, 4))
        Integer.TryParse(Me.TextBox_42.Text, TempGridArray(4, 5))
        Integer.TryParse(Me.TextBox_43.Text, TempGridArray(4, 6))
        Integer.TryParse(Me.TextBox_44.Text, TempGridArray(4, 7))
        Integer.TryParse(Me.TextBox_45.Text, TempGridArray(4, 8))
        Integer.TryParse(Me.TextBox_46.Text, TempGridArray(5, 0))
        Integer.TryParse(Me.TextBox_47.Text, TempGridArray(5, 1))
        Integer.TryParse(Me.TextBox_48.Text, TempGridArray(5, 2))
        Integer.TryParse(Me.TextBox_49.Text, TempGridArray(5, 3))
        Integer.TryParse(Me.TextBox_50.Text, TempGridArray(5, 4))
        Integer.TryParse(Me.TextBox_51.Text, TempGridArray(5, 5))
        Integer.TryParse(Me.TextBox_52.Text, TempGridArray(5, 6))
        Integer.TryParse(Me.TextBox_53.Text, TempGridArray(5, 7))
        Integer.TryParse(Me.TextBox_54.Text, TempGridArray(5, 8))
        Integer.TryParse(Me.TextBox_55.Text, TempGridArray(6, 0))
        Integer.TryParse(Me.TextBox_56.Text, TempGridArray(6, 1))
        Integer.TryParse(Me.TextBox_57.Text, TempGridArray(6, 2))
        Integer.TryParse(Me.TextBox_58.Text, TempGridArray(6, 3))
        Integer.TryParse(Me.TextBox_59.Text, TempGridArray(6, 4))
        Integer.TryParse(Me.TextBox_60.Text, TempGridArray(6, 5))
        Integer.TryParse(Me.TextBox_61.Text, TempGridArray(6, 6))
        Integer.TryParse(Me.TextBox_62.Text, TempGridArray(6, 7))
        Integer.TryParse(Me.TextBox_63.Text, TempGridArray(6, 8))
        Integer.TryParse(Me.TextBox_64.Text, TempGridArray(7, 0))
        Integer.TryParse(Me.TextBox_65.Text, TempGridArray(7, 1))
        Integer.TryParse(Me.TextBox_66.Text, TempGridArray(7, 2))
        Integer.TryParse(Me.TextBox_67.Text, TempGridArray(7, 3))
        Integer.TryParse(Me.TextBox_68.Text, TempGridArray(7, 4))
        Integer.TryParse(Me.TextBox_69.Text, TempGridArray(7, 5))
        Integer.TryParse(Me.TextBox_70.Text, TempGridArray(7, 6))
        Integer.TryParse(Me.TextBox_71.Text, TempGridArray(7, 7))
        Integer.TryParse(Me.TextBox_72.Text, TempGridArray(7, 8))
        Integer.TryParse(Me.TextBox_73.Text, TempGridArray(8, 0))
        Integer.TryParse(Me.TextBox_74.Text, TempGridArray(8, 1))
        Integer.TryParse(Me.TextBox_75.Text, TempGridArray(8, 2))
        Integer.TryParse(Me.TextBox_76.Text, TempGridArray(8, 3))
        Integer.TryParse(Me.TextBox_77.Text, TempGridArray(8, 4))
        Integer.TryParse(Me.TextBox_78.Text, TempGridArray(8, 5))
        Integer.TryParse(Me.TextBox_79.Text, TempGridArray(8, 6))
        Integer.TryParse(Me.TextBox_80.Text, TempGridArray(8, 7))
        Integer.TryParse(Me.TextBox_81.Text, TempGridArray(8, 8))

        Return TempGridArray
    End Function

    Private Sub SetTextBoxesData(Grid(,) As Integer)
        Me.TextBox_1.Text = Grid(0, 0).ToString
        Me.TextBox_2.Text = Grid(0, 1).ToString
        Me.TextBox_3.Text = Grid(0, 2).ToString
        Me.TextBox_4.Text = Grid(0, 3).ToString
        Me.TextBox_5.Text = Grid(0, 4).ToString
        Me.TextBox_6.Text = Grid(0, 5).ToString
        Me.TextBox_7.Text = Grid(0, 6).ToString
        Me.TextBox_8.Text = Grid(0, 7).ToString
        Me.TextBox_9.Text = Grid(0, 8).ToString
        Me.TextBox_10.Text = Grid(1, 0).ToString
        Me.TextBox_11.Text = Grid(1, 1).ToString
        Me.TextBox_12.Text = Grid(1, 2).ToString
        Me.TextBox_13.Text = Grid(1, 3).ToString
        Me.TextBox_14.Text = Grid(1, 4).ToString
        Me.TextBox_15.Text = Grid(1, 5).ToString
        Me.TextBox_16.Text = Grid(1, 6).ToString
        Me.TextBox_17.Text = Grid(1, 7).ToString
        Me.TextBox_18.Text = Grid(1, 8).ToString
        Me.TextBox_19.Text = Grid(2, 0).ToString
        Me.TextBox_20.Text = Grid(2, 1).ToString
        Me.TextBox_21.Text = Grid(2, 2).ToString
        Me.TextBox_22.Text = Grid(2, 3).ToString
        Me.TextBox_23.Text = Grid(2, 4).ToString
        Me.TextBox_24.Text = Grid(2, 5).ToString
        Me.TextBox_25.Text = Grid(2, 6).ToString
        Me.TextBox_26.Text = Grid(2, 7).ToString
        Me.TextBox_27.Text = Grid(2, 8).ToString
        Me.TextBox_28.Text = Grid(3, 0).ToString
        Me.TextBox_29.Text = Grid(3, 1).ToString
        Me.TextBox_30.Text = Grid(3, 2).ToString
        Me.TextBox_31.Text = Grid(3, 3).ToString
        Me.TextBox_32.Text = Grid(3, 4).ToString
        Me.TextBox_33.Text = Grid(3, 5).ToString
        Me.TextBox_34.Text = Grid(3, 6).ToString
        Me.TextBox_35.Text = Grid(3, 7).ToString
        Me.TextBox_36.Text = Grid(3, 8).ToString
        Me.TextBox_37.Text = Grid(4, 0).ToString
        Me.TextBox_38.Text = Grid(4, 1).ToString
        Me.TextBox_39.Text = Grid(4, 2).ToString
        Me.TextBox_40.Text = Grid(4, 3).ToString
        Me.TextBox_41.Text = Grid(4, 4).ToString
        Me.TextBox_42.Text = Grid(4, 5).ToString
        Me.TextBox_43.Text = Grid(4, 6).ToString
        Me.TextBox_44.Text = Grid(4, 7).ToString
        Me.TextBox_45.Text = Grid(4, 8).ToString
        Me.TextBox_46.Text = Grid(5, 0).ToString
        Me.TextBox_47.Text = Grid(5, 1).ToString
        Me.TextBox_48.Text = Grid(5, 2).ToString
        Me.TextBox_49.Text = Grid(5, 3).ToString
        Me.TextBox_50.Text = Grid(5, 4).ToString
        Me.TextBox_51.Text = Grid(5, 5).ToString
        Me.TextBox_52.Text = Grid(5, 6).ToString
        Me.TextBox_53.Text = Grid(5, 7).ToString
        Me.TextBox_54.Text = Grid(5, 8).ToString
        Me.TextBox_55.Text = Grid(6, 0).ToString
        Me.TextBox_56.Text = Grid(6, 1).ToString
        Me.TextBox_57.Text = Grid(6, 2).ToString
        Me.TextBox_58.Text = Grid(6, 3).ToString
        Me.TextBox_59.Text = Grid(6, 4).ToString
        Me.TextBox_60.Text = Grid(6, 5).ToString
        Me.TextBox_61.Text = Grid(6, 6).ToString
        Me.TextBox_62.Text = Grid(6, 7).ToString
        Me.TextBox_63.Text = Grid(6, 8).ToString
        Me.TextBox_64.Text = Grid(7, 0).ToString
        Me.TextBox_65.Text = Grid(7, 1).ToString
        Me.TextBox_66.Text = Grid(7, 2).ToString
        Me.TextBox_67.Text = Grid(7, 3).ToString
        Me.TextBox_68.Text = Grid(7, 4).ToString
        Me.TextBox_69.Text = Grid(7, 5).ToString
        Me.TextBox_70.Text = Grid(7, 6).ToString
        Me.TextBox_71.Text = Grid(7, 7).ToString
        Me.TextBox_72.Text = Grid(7, 8).ToString
        Me.TextBox_73.Text = Grid(8, 0).ToString
        Me.TextBox_74.Text = Grid(8, 1).ToString
        Me.TextBox_75.Text = Grid(8, 2).ToString
        Me.TextBox_76.Text = Grid(8, 3).ToString
        Me.TextBox_77.Text = Grid(8, 4).ToString
        Me.TextBox_78.Text = Grid(8, 5).ToString
        Me.TextBox_79.Text = Grid(8, 6).ToString
        Me.TextBox_80.Text = Grid(8, 7).ToString
        Me.TextBox_81.Text = Grid(8, 8).ToString

    End Sub
End Class
