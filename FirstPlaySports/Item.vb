'Program:       FirstPlaySports
'Developer:     Jason A. Frye
'Date:          10/2/11
'Purpose:       This object will load the sports equipment inventory from a text file, 
'               allow the user to display, edit, add, or remove entries, and save the information 
'               back to the same text file as the user exits. 

Option Strict On
Option Explicit On
Imports System.IO

Public Class Item
    'the Item class is the encapsulation of one 'Item' from the inventory. 

    Public Property ID As String
    Public Property Description As String
    Public Property Daily As Double
    Public Property Weekly As Double
    Public Property Monthly As Double
    Public Property Qty As Integer

End Class
