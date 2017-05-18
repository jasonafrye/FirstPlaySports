'Program:       FirstPlaySports
'Developer:     Jason A. Frye
'Date:          10/2/11
'Purpose:       This object will load the sports equipment inventory from a text file, 
'               allow the user to display, edit, add, or remove entries, and save the information 
'               back to the same text file as the user exits. 

Option Strict On
Option Explicit On
Imports System.IO

Public Class frmMain

#Region "Control Events"

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'This subroutine calls then Inventory.LoadInventory fuction to retrieve the boolean flag. 
        'If unsuccessful, the user is presented with an error and the program closes, 
        'else calls the PopulateForm() subroutine

        'Declare Flag
        Dim blnLoad As Boolean = False

        'Call business tier function
        blnLoad = Inventory.LoadInventory()

        'test flag
        If blnLoad = False Then
            'present erro and terminate application
            MsgBox("Unable to Load Inventory!", , "Error!")
            Me.Close()
        Else
            'call population subroutine
            PopulateForm()
        End If

    End Sub

    Private Sub btnClose_Click(sender As Object, e As System.EventArgs) Handles btnClose.Click
        'This subroutine calls the Inventory.SaveInventory subroutine to retrieve the boolean flag. 
        'If successful the application is terminated, else the user is presented with an error message

        'declare flag 
        Dim blnSave As Boolean = False

        'call business tier function
        blnSave = Inventory.SaveInventory

        'test flag
        If blnSave = True Then
            'terminate program
            Me.Close()
        Else
            'present error message
            MsgBox("There is a problem saving your file, try again later", , "Uh Oh!")
        End If

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
        'This subroutine verifies that all fields have data present, 
        'if so: instantiates new Item object and sends object to Inventory.AddItem 
        'subroutine and calls PopulateForm subroutine. 
        'If not: provides an errorand sets focus to dropdown list 

        'test that ID number is entered
        If Not cboID.Text = Nothing Then
            ErrorProvider1.Clear()

            'test that other textboxes have data present
            If txtDaily.Text = Nothing Or txtWeekly.Text = Nothing Or txtMonthly.Text = Nothing Or txtQty.Text = Nothing Or txtDesc.Text = Nothing Then

                'if not, set error
                ErrorProvider1.SetError(btnAdd, "Please fill in all fields")

            Else
                'if so, clear error
                ErrorProvider1.Clear()

                'isntantiate new Item object
                Dim newItem As New Item
                'set properties based on textbox values
                With newItem
                    .ID = cboID.Text.ToString
                    .Description = txtDesc.Text
                    .Daily = CDbl(txtDaily.Text)
                    .Weekly = CDbl(txtWeekly.Text)
                    .Monthly = CDbl(txtMonthly.Text)
                    .Qty = CInt(txtQty.Text)
                End With

                'Pass to subroutine in business tier
                Inventory.AddItem(newItem)
                'recall PopulateForm subroutinte
                PopulateForm()

            End If

        Else
            'if ID is not present, set error to try again. 
            ErrorProvider1.SetError(cboID, "Please Select an Inventory Item")
            cboID.Focus()

        End If

    End Sub

    Private Sub btnDisplay_Click(sender As Object, e As System.EventArgs) Handles btnDisplay.Click
        'This subroutine verifies that the dropdownlist item has been selected, 
        'and then passes that key to the Inventory.FindItem function to retrieve item from the Dictionary. 
        'Then that item's attributes are assigned to the appropriate textboxes and is presented to the user. 
        'If an ID is not selected, the user is presented with an error message and focus is reset

        'test that ID has been selected
        If cboID.SelectedIndex >= 0 Then

            'if successfull, retrieve string data and declare new Item object
            Dim keyID As String = cboID.SelectedItem.ToString
            Dim SelectedItem As New Item

            'send keyID to business tier function to retrieve item from dictioanry
            SelectedItem = Inventory.FindItem(keyID)

            'allocate item properties to the appropriate textboxes for display
            txtDesc.Text = SelectedItem.Description.ToString
            txtDaily.Text = SelectedItem.Daily.ToString
            txtWeekly.Text = SelectedItem.Weekly.ToString
            txtMonthly.Text = SelectedItem.Monthly.ToString
            txtQty.Text = SelectedItem.Qty.ToString

        Else
            'if unsuccessfull, set error and reset focus to dropdownlist
            ErrorProvider1.SetError(cboID, "Please Select an Inventory Item")
            cboID.Focus()
        End If

    End Sub

    Private Sub btnRemove_Click(sender As Object, e As System.EventArgs) Handles btnRemove.Click
        'This subroutine verifies that the dropdownlist item has been selected, 
        'if so the user is presented with a confirmation dialogue and the item key is sent  
        'to the Inventory.RemoveItem subroutine  and the form is repopulated. 


        'test that dropdownlist item has been selected
        If cboID.SelectedIndex >= 0 Then
            'if successful, declare answer variable and prompt the user
            Dim Answer As Microsoft.VisualBasic.MsgBoxResult
            Answer = MsgBox("Do you wish to remove this inventory item? This cannot be undone.", MsgBoxStyle.YesNo, "Are You Sure?")

            If Answer = vbYes Then
                'if successful, retrieve string value of dropdownlist
                Dim keyID As String = cboID.SelectedItem.ToString
                'pass keyID to business tier function
                Inventory.RemoveItem(keyID)
                'repopulate form
                PopulateForm()
            End If
        Else
            'if unsuccessful, set error and reset focus. 
            ErrorProvider1.SetError(cboID, "Please Select an Inventory Item")
            cboID.Focus()
        End If
    End Sub

#End Region

#Region "Validation Events"

    Private Sub txtDaily_TextChanged(sender As Object, e As System.EventArgs) Handles txtDaily.TextChanged
        'this subroutine tests that every character entered into the textbox can be converted to the appropriate datatype. 
        'If not, the error is set, the box is cleared, and the focus is reset

        Dim dblTest As Double
        Try
            ErrorProvider1.Clear()
            dblTest = Convert.ToDouble(txtDaily.Text)
        Catch ex As Exception
            ErrorProvider1.SetError(txtDaily, "Please enter a positive numeric value")
            txtDaily.Clear()
            txtDaily.Focus()
        End Try
    End Sub

    Private Sub txtMonthly_TextChanged(sender As Object, e As System.EventArgs) Handles txtMonthly.TextChanged
        'this subroutine tests that every character entered into the textbox can be converted to the appropriate datatype. 
        'If not, the error is set, the box is cleared, and the focus is reset

        Dim dblTest As Double
        Try
            ErrorProvider1.Clear()
            dblTest = Convert.ToDouble(txtMonthly.Text)
        Catch ex As Exception
            ErrorProvider1.SetError(txtMonthly, "Please enter a positive numeric value")
            txtMonthly.Clear()
            txtMonthly.Focus()
        End Try
    End Sub

    Private Sub txtWeekly_TextChanged(sender As Object, e As System.EventArgs) Handles txtWeekly.TextChanged
        'this subroutine tests that every character entered into the textbox can be converted to the appropriate datatype. 
        'If not, the error is set, the box is cleared, and the focus is reset

        Dim dblTest As Double
        Try
            ErrorProvider1.Clear()
            dblTest = Convert.ToDouble(txtWeekly.Text)
        Catch ex As Exception
            ErrorProvider1.SetError(txtWeekly, "Please enter a positive numeric value")
            txtWeekly.Clear()
            txtWeekly.Focus()
        End Try
    End Sub

    Private Sub txtQty_TextChanged(sender As Object, e As System.EventArgs) Handles txtQty.TextChanged
        'this subroutine tests that every character entered into the textbox can be converted to the appropriate datatype. 
        'If not, the error is set, the box is cleared, and the focus is reset

        Dim intTest As Integer
        Try
            ErrorProvider1.Clear()
            intTest = Convert.ToInt32(txtDaily.Text)
        Catch ex As Exception
            ErrorProvider1.SetError(txtDaily, "Please enter a positive numeric value")
            txtDaily.Clear()
            txtDaily.Focus()
        End Try
    End Sub

#End Region

#Region "Custom Procedures"

    Private Sub PopulateForm()
        'The subroutine removes data from the entire form, 
        'calls for the latest version of the dictionary and 
        'populates the dropdownlist with all available KeyIDs

        'Clear all form controls
        cboID.Items.Clear()
        txtDaily.Text = ""
        txtDesc.Text = ""
        txtMonthly.Text = ""
        txtQty.Text = ""
        txtWeekly.Text = ""
        cboID.Text = ""
        ErrorProvider1.Clear()

        'retrieve latest version of the inventory dictionary
        Dim inventoryList As Dictionary(Of String, Item)
        inventoryList = CType(Inventory.PopulateForm, Global.System.Collections.Generic.Dictionary(Of String, Global.FirstPlaySports.Item))

        'loops through dictionary and adds all ID properties to the dropdown list
        For Each entry In inventoryList
            cboID.Items.Add(entry.Key)
        Next

    End Sub

#End Region

End Class
