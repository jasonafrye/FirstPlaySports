'Program:       FirstPlaySports
'Developer:     Jason A. Frye
'Date:          10/2/11
'Purpose:       This object will load the sports equipment inventory from a text file, 
'               allow the user to display, edit, add, or remove entries, and save the information 
'               back to the same text file as the user exits. 

Option Strict On
Option Explicit On
Imports System.IO


Public Class Inventory_File

    Public Shared Function LoadData() As IDictionary
        'This function reads the data from the textfile, creates the distionary, 
        'populates it with objects, and returns entire dictionary to calling procedure

        'declare file location and dictioanry
        Dim strFile As String = "..\inventory.txt"
        Dim inventoryList As New Dictionary(Of String, Item)

        'verify file exists
        If IO.File.Exists(strFile) Then

            'if successful, declare stream reader to open the file
            Dim objReader As New StreamReader(strFile)

            'loop through file to retrieve data
            While Not objReader.EndOfStream
                'instantiate new Item object
                Dim newItem As New Item
                'assign attributes
                With newItem
                    .ID = objReader.ReadLine
                    .Description = objReader.ReadLine
                    .Daily = CDbl(objReader.ReadLine)
                    .Weekly = CDbl(objReader.ReadLine)
                    .Monthly = CDbl(objReader.ReadLine)
                    .Qty = CInt(objReader.ReadLine)
                End With
                'add new item to dictioanry
                inventoryList.Add(newItem.ID, newItem)
            End While
            'close streamreader

            objReader.Close()
        End If

        Return inventoryList
    End Function

    Public Shared Function SaveData(ByVal NewInventory As IDictionary) As Boolean
        'this function opens the text file, pulls each object from the dictionary and 
        'writes it back to the text file and returns boolean flag to calling procedure. 

        'declare flag and file location
        Dim blnSave As Boolean = False
        Dim strFile As String = "..\inventory.txt"

        'test that file exists
        If IO.File.Exists(strFile) Then

            'open StreamWriter, declare strings for object attributse, and copy current dictionary for saving
            Dim objWriter As New StreamWriter(strFile)
            Dim saveID As String
            Dim saveDesc As String
            Dim saveDaily As String
            Dim saveWeekly As String
            Dim saveMonthly As String
            Dim saveQty As String
            Dim SaveInventory As New Dictionary(Of String, Item)

            SaveInventory = CType(NewInventory, Global.System.Collections.Generic.Dictionary(Of String, Global.FirstPlaySports.Item))

            'loop through the dictioanry for each item
            For Each entry In SaveInventory
                'declare new Item
                Dim saveItem As New Item
                'assign object from dictionary to new item
                saveItem = SaveInventory.Item(entry.Key)

                'assign item attributes to string values for writing
                saveID = saveItem.ID.ToString
                saveDesc = saveItem.Description.ToString
                saveDaily = saveItem.Daily.ToString
                saveWeekly = saveItem.Weekly.ToString
                saveMonthly = saveItem.Monthly.ToString
                saveQty = saveItem.Qty.ToString

                'Write said string values to textfile
                objWriter.WriteLine(saveID)
                objWriter.WriteLine(saveDesc)
                objWriter.WriteLine(saveDaily)
                objWriter.WriteLine(saveWeekly)
                objWriter.WriteLine(saveMonthly)
                objWriter.WriteLine(saveQty)
            Next
            'close StreamWriter and set confirmation flag to True
            objWriter.Close()
            blnSave = True
        End If

        Return blnSave
    End Function

End Class

