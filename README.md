<div align="center">

## DBGrid Dropdown Listbox


</div>

### Description

Ever wondered what the button_click event was for in a DBGrid? Well, this is it! You have to populate a listbox control and display that with the selection information. This will simulate a dropdown box within the dbgrid. This is an excelent way to input specific information into the dbgrid.
 
### More Info
 
Nont

Lose the ability to use the down arrow key in the column with the button.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jason J\. Martin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-j-martin.md)
**Level**          |Unknown
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jason-j-martin-dbgrid-dropdown-listbox__1-1377/archive/master.zip)

### API Declarations

```
Global Declarations within form:
Dim intColIdx As Integer 'This will contain the index for the current cell in the dbgrid
Dim blnListShow As Boolean 'is the list showing or not
Dim intKeyCode As Integer 'the last key pressed within the dbgrid
```


### Source Code

```
Create a form with a dbgrid(DBGrid1), and a listbox(List1). Populate the listbox with the choices you need the user to select from. Set the visible property on the listbox to false. Set the button property on one of the DBGrid columns to true. This example is using column 2. If you want to limit the input to the DBGrid to just the items in the listbox, set the enabled property to false, otherwise, users can type in their own data.
Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
  Dim intTop As Integer 'used for positioning the list box for display.
  intColIdx = ColIndex 'this is the column of the dbgrid you are in
  If blnListShow = False Then 'if the list is not showing then...
    blnListShow = True
    List1.Left = DBGrid1.Columns(ColIndex).Left + 250 'you may have to play
                                            'with this a little to get it
                                            'positioned just right.
    intTop = DBGrid1.Top + (DBGrid1.RowHeight * (DBGrid1.Row + 2))
    List1.Top = intTop 'position the list box just below the row you are in
    List1.Width = DBGrid1.Columns(ColIndex).Width + 15 'setting the width of
                                               'the listbox to display
                                               'within the column
                                               ' width
    List1.Visible = True 'show the listbox
    List1.SetFocus
  Else 'if the list is shown, hide it
    blnListShow = False
    List1.Visible = False
  End If
End Sub
Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  'This is to display the list when the user presses the down arrow key.
  'This makes it easier to make a selection during data entry. The user
  'doesn't have to go to the mouse to click the button.
  If DBGrid1.Col = 2 Then 'change the number here to your appropriate column
                      'that has the button, other wise you will display the
                      ' listbox on the wrong column
    If KeyCode = vbKeyDown Then
      Call DBGrid1_ButtonClick(DBGrid1.Col)
    End If
  End If
End Sub
Private Sub Form_Click()
  'hide the listbox if the user clicks elsewhere
  List1.Visible = False 'hide the list
End Sub
Private Sub Form_Load()
  blnListShow = False 'initialize the variable
End Sub
Private Sub Form_Resize()
  'hide the list if they resize the form
  List1.Visible = False 'hide the list
End Sub
Private Sub List1_Click()
  'insert the selected list item into the dbgrid, and hide the listbox
  If intKeyCode <> vbKeyUp And intKeyCode <> vbKeyDown Then
    DBGrid1.Columns(intColIdx).Text = List1.Text 'set the value of the dbgrid
    List1.Visible = False 'hide the list
  End If
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
  'handle the keyboard events
  intKeyCode = KeyCode
  If intKeyCode = vbKeyReturn Then
    DBGrid1.Columns(intColIdx).Text = List1.Text 'set the value of the dbgrid
    List1.Visible = False 'hide the list
  Else
    If intKeyCode = vbKeyEscape Then
      List1.Visible = False
    End If
  End If
End Sub
Private Sub List1_LostFocus()
'hide the list if you lose focus
  blnListShow = False
  List1.Visible = False
End Sub
```

