'Program:       FirstPlaySports
'Developer:     Jason A. Frye
'Date:          10/2/11
'Purpose:       This object will load the sports equipment inventory from a text file, 
'               allow the user to display, edit, add, or remove entries, and save the information 
'               back to the same text file as the user exits. 

Option Strict On
Option Explicit On
Imports System.IO

Public Class Inventory

    'declare dictionary to be used throughout the application
    Private Shared inventoryList As New Dictionary(Of String, Item)

    Public Shared Sub AddItem(ByVal newItem As Item)
        'receives the new Item object from the form and adds it to the current dictionary, 

        inventoryList.Item(newItem.ID) = newItem

    End Sub

    Public Shared Sub RemoveItem(ByVal keyID As String)
        'receives the keyID from the form and removes that item from the current dictionary

        inventoryList.Remove(keyID)

    End Sub

    Public Shared Function PopulateForm() As IDictionary
        'returns the most current version of the inventory list to the calling procedure

        Return inventoryList

    End Function

    Public Shared Function LoadInventory() As Boolean
        'This function tries to call the Inventory_File.LoadData function 
        'from the persistence tier and returns the boolean flag value to the calling procedure

        'declare flagg
        Dim blnLoad As Boolean = False

        Try
            'call persistence tier function and set flag
            inventoryList = CType(Inventory_File.LoadData, Global.System.Collections.Generic.Dictionary(Of String, Global.FirstPlaySports.Item))
            blnLoad = True
        Catch ex As Exception
            blnLoad = False
        End Try

        Return blnLoad

    End Function

    Public Shared Function FindItem(ByVal keyID As String) As Item
        'this function receives the keyID from the form, finds the appropriate 
        'dictioanry entry and sends it back to the calling procedure

        'instantiate new object
        Dim Entry As Item
        Entry = inventoryList.Item(keyID)

        Return Entry

    End Function

    Public Shared Function SaveInventory() As Boolean
        'This function tries to call the Inventory_File.SaveData function 
        'from the persistence tier and returns the boolean flag value to the calling procedure

        'declare flag
        Dim blnSave As Boolean = False

        'call function from persistence tier
        blnSave = Inventory_File.SaveData(inventoryList)

        Return blnSave

    End Function

End Class
