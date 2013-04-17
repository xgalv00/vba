VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} copyMineUF 
   Caption         =   "Консолидация юрлица"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   OleObjectBlob   =   "copyMineUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "copyMineUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim uf_constValColl As Collection

Private Sub UserForm_Initialize()
'''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''
'userforms's constant value initialization

    'don't change values of this collection in your code. Should be replaced by some immutable type, but i don't
    'know what I should use in excel
    Set uf_constValColl = New Collection
    
    uf_constValColl.Add "mineMan", "mine_management_prefix"
    uf_constValColl.Add "mine", "mine_prefix"
    uf_constValColl.Add "CmBx", "cmbx_postfix"
    uf_constValColl.Add "Lbl", "lbl_postfix"
    
End Sub

Private Sub cancelBtn_Click()
    '@proc from Controller
    Call Controller.unloadCopyMineUF
    
End Sub

Private Sub chooseSrcBtn_Click()
    ' Display full path and name of the files
    Dim tmpStr As String
    '@proc from Controller
    tmpStr = Controller.proccesFileSelection
    If tmpStr = "" Then
        Exit Sub
    End If
    srcNameLbl.Caption = tmpStr
    
    If Not mineManCmBx.Enabled Then
        Call enableCmbx(uf_constValColl("mine_management_prefix"))
    End If
    
End Sub

Private Sub copyBtn_Click()
    '@proc from Controller
    Call Controller.copyBtnClicked
    '@proc from Controller
    Call Controller.unloadCopyMineUF
End Sub

Private Sub copyStyleChkBx_Click()
    If mineManCmBx.Text <> "" Then 'do nothing if first combobox is empty
        Dim cmbxName As String
        
        cmbxName = uf_constValColl("mine_prefix") & uf_constValColl("cmbx_postfix")
        
        If copyStyleChkBx Then
            togle_black_red (uf_constValColl("mine_prefix"))
            Call enableCmbx(uf_constValColl("mine_prefix"))
        Else
            If copyMineUF.Controls(cmbxName).Enabled Then
                If copyMineUF.Controls(cmbxName).Text = "" Then
                    togle_black_red (uf_constValColl("mine_prefix"))
                End If
                disable_cmbx (uf_constValColl("mine_prefix"))
            End If
        End If
    End If
End Sub

Private Sub mineCmBx_Change()
    If Not Controller.techChange Then
    '@proc from Controller
        Call Controller.generalFiltering
        Call enable_copyBtn
        togle_black_red (uf_constValColl("mine_prefix"))
    End If
End Sub

Private Sub mineManCmBx_Change()
    If Not Controller.techChange Then
    '@proc from Controller
        Call Controller.generalFiltering
        togle_black_red (uf_constValColl("mine_management_prefix"))

        If copyStyleChkBx Then
            'copyMineUF.mineCmBx.Enabled = True
            togle_black_red (uf_constValColl("mine_prefix"))
            'copyMineUF.mineLbl.ForeColor = vbRed
            'copyMineUF.mineCmBx.RowSource = Controller.computerRowSource("mineCmBx")
            Call enableCmbx(uf_constValColl("mine_prefix"))
        Else
            Call enable_copyBtn
        End If
    End If
End Sub

Sub set_file_name_label(FileName As String)
    srcNameLbl.Caption = FileName
End Sub

Private Sub enableCmbx(cmbxName As String)
    Dim tmpStr As String

    tmpStr = cmbxName & uf_constValColl("cmbx_postfix")
    
    copyMineUF.Controls(tmpStr).Enabled = True
    copyMineUF.Controls(tmpStr).RowSource = Controller.computerRowSource(tmpStr)

End Sub

Private Sub enable_copyBtn()
    copyMineUF.copyBtn.Enabled = True
End Sub

Private Sub togle_black_red(lblPrefix As String)
    Dim tmpStr As String
    
    tmpStr = lblPrefix & uf_constValColl("lbl_postfix")
    If copyMineUF.Controls(tmpStr).ForeColor = vbRed Then
        copyMineUF.Controls(tmpStr).ForeColor = vbBlack
    Else
        copyMineUF.Controls(tmpStr).ForeColor = vbRed
    End If
End Sub

Private Sub disable_cmbx(cmbx_prefix As String)

    Dim tmpStr As String

    tmpStr = cmbx_prefix & uf_constValColl("cmbx_postfix")
    
    copyMineUF.Controls(tmpStr).Enabled = False
    copyMineUF.Controls(tmpStr).Text = ""

End Sub

