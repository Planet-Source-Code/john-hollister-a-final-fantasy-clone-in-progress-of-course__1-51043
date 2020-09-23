Attribute VB_Name = "storeMod"
Global storeInv(0 To 1, 0 To 50, 0 To 25) As Variant
Global storeIDs(0 To 9, 0 To 9, 0 To 19, 0 To 19) As Integer
'0 to 1 is for name/price
'0 to 50 is for # of stores in game
'0 to 25 is for # of items avail in store

'*********************************************
'* set the inventory for a specific store here
'* the sID (storeID) is stored in an array in RPGvars
'*********************************************

Public Sub getStore(sID As Integer)

'must figure out how to make stores unique
    If sID = 0 Then
        storeInv(0, sID, 0) = "Potion"
        storeInv(1, sID, 0) = 50
        storeInv(0, sID, 1) = "Ether"
        storeInv(1, sID, 1) = 100
        storeInv(0, sID, 2) = "Phoenix"
        storeInv(1, sID, 2) = 100
        storeInv(0, sID, 3) = "Tent"
        storeInv(1, sID, 3) = 400
    End If


End Sub

