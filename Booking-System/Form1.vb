

Imports System.Configuration
Imports System.Data.OleDb

Public Class Form1

    Dim availableIcon As New System.Drawing.Bitmap(My.Resources.availiable)
    Dim provisionalIcon As New System.Drawing.Bitmap(My.Resources.provisional)
    Dim bookedIcon As New System.Drawing.Bitmap(My.Resources.booked)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each c As Control In Me.Controls
            If TypeOf (c) Is PictureBox Then
                Dim pictureBox As PictureBox = CType(c, PictureBox)
                pictureBox.SizeMode = PictureBoxSizeMode.StretchImage
                pictureBox.Width = 48 ' Width of the PictureBox
                pictureBox.Height = 40 ' Height of the PictureBox
                pictureBox.Image = availableIcon
                pictureBox.Tag = "available" ' Set initial state
                AddHandler pictureBox.Click, AddressOf PictureBox_Click
            End If
        Next

        Dim stSQL As String = "SELECT BookingID, CustomerID, Seat From tblBookings"
        Dim stConString As String
        stConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\malum\Documents\BookingDatabase.accdb"

        Dim conBooking As New OleDbConnection

        conBooking.ConnectionString = stConString
        conBooking.Open()

        Dim cmdSelectBookings As New OleDbCommand
        cmdSelectBookings.CommandText = stSQL
        cmdSelectBookings.Connection = conBooking

        Dim dsBookings As New DataSet
        Dim daBookings As New OleDbDataAdapter(cmdSelectBookings)

        daBookings.Fill(dsBookings, "Bookings")

        conBooking.Close()

        Dim stOut As String
        Dim t1 As DataTable = dsBookings.Tables("Bookings")
        Dim row As DataRow

        For Each row In t1.Rows
            stOut = stOut & row(0) & " " & row(1) & " " & row(2) & vbNewLine
            CType(Controls("PictureBox" & row(2)), PictureBox).Image = bookedIcon
        Next


    End Sub

    Private Sub PictureBox_Click(sender As Object, e As EventArgs)
        Dim pictureBox As PictureBox = TryCast(sender, PictureBox)

        If pictureBox.Tag.ToString() = "available" Then
            pictureBox.Image = provisionalIcon
            pictureBox.Tag = "provisional"
        ElseIf pictureBox.Tag.ToString() = "provisional" Then
            pictureBox.Image = availableIcon
            pictureBox.Tag = "available"
        End If
    End Sub

    Private Sub btnContinue_Click(sender As Object, e As EventArgs) Handles btnContinue.Click

        Try
            Dim c As Control
            Dim bSelected As Boolean

            For Each c In Me.Controls
                If TypeOf (c) Is PictureBox Then
                    If CType(c, PictureBox).Image Is provisionalIcon Then
                        bSelected = True
                        Exit For
                    End If
                    AddHandler c.Click, AddressOf PictureBox_Click
                End If
            Next
            If Not bSelected Then
                MsgBox("Please select at least one seat to book")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



        Dim stSQLInsert As String = "INSERT INTO tblBookings (CustomerID, Seat) VALUES('AC001', 78)"
        Dim stConString As String
        stConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\malum\Documents\BookingDatabase.accdb"

        Dim conBooking As New OleDbConnection

        conBooking.ConnectionString = stConString
        conBooking.Open()

        Dim cmdMakeBookings As New OleDbCommand

        cmdMakeBookings.Connection = conBooking



        Dim iSeatNum As Integer

        Try

            For Each c In Me.Controls
                If TypeOf (c) Is PictureBox Then
                    If CType(c, PictureBox).Image Is provisionalIcon Then
                        iSeatNum = Mid(CType(c, PictureBox).Name, 11)

                        stSQLInsert = "INSERT INTO tblBookings (CustomerID, Seat) VALUES('" & Me.txtCustomer.Text & "', " & iSeatNum & ")"
                        cmdMakeBookings.CommandText = stSQLInsert
                        cmdMakeBookings.ExecuteNonQuery()
                    End If

                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub
End Class





