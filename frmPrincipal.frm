VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Mega Sena"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "|^N�mero|^Vezes sorteadas "
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gerar n�meros mais sorteados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carregar resultador para o Banco de Dados"
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Banco As Database
Dim Result As Recordset
Dim Tempor�rio As Recordset

Dim ArqRelat�rio As String
Dim strLinha As String
Dim N�mero As String
Dim D1 As String
Dim D2 As String
Dim D3 As String
Dim D4 As String
Dim D5 As String
Dim D6 As String
Dim Deleta As String

Private Sub Command1_Click()
    
    Deleta = "Delete * from Resultados"
    Banco.Execute Deleta
    
    ArqRelat�rio = App.Path & "\Resultados Mega Sena.txt"
    If Dir(ArqRelat�rio) = "" Then
        MsgBox "Arquivo n�o encontrado!", vbInformation
    Else
        Open ArqRelat�rio For Input As #1
        While Not EOF(1)
            Line Input #1, strLinha
            N�mero = Trim(Mid(strLinha, 1, (InStr(1, strLinha, vbTab) - 1)))
            If N�mero < 10 Then
                D1 = Trim(Mid(strLinha, 16, ((InStr(16, strLinha, vbTab) - 1) - 16)))
                D2 = Trim(Mid(strLinha, 20, ((InStr(20, strLinha, vbTab) - 1) - 20)))
                D3 = Trim(Mid(strLinha, 24, ((InStr(24, strLinha, vbTab) - 1) - 24)))
                D4 = Trim(Mid(strLinha, 28, ((InStr(28, strLinha, vbTab) - 1) - 28)))
                D5 = Trim(Mid(strLinha, 32, ((InStr(32, strLinha, vbTab) - 1) - 32)))
                D6 = Trim(Mid(strLinha, 36, ((InStr(36, strLinha, vbTab) - 1) - 36)))
            ElseIf N�mero >= 10 And N�mero < 99 Then
                D1 = Trim(Mid(strLinha, 17, ((InStr(17, strLinha, vbTab) - 1) - 17)))
                D2 = Trim(Mid(strLinha, 21, ((InStr(21, strLinha, vbTab) - 1) - 21)))
                D3 = Trim(Mid(strLinha, 25, ((InStr(25, strLinha, vbTab) - 1) - 25)))
                D4 = Trim(Mid(strLinha, 29, ((InStr(29, strLinha, vbTab) - 1) - 29)))
                D5 = Trim(Mid(strLinha, 33, ((InStr(33, strLinha, vbTab) - 1) - 33)))
                D6 = Trim(Mid(strLinha, 37, ((InStr(37, strLinha, vbTab) - 1) - 37)))
            ElseIf N�mero >= 100 Then
                D1 = Trim(Mid(strLinha, 18, ((InStr(18, strLinha, vbTab) - 1) - 18)))
                D2 = Trim(Mid(strLinha, 22, ((InStr(22, strLinha, vbTab) - 1) - 22)))
                D3 = Trim(Mid(strLinha, 26, ((InStr(26, strLinha, vbTab) - 1) - 26)))
                D4 = Trim(Mid(strLinha, 30, ((InStr(30, strLinha, vbTab) - 1) - 30)))
                D5 = Trim(Mid(strLinha, 34, ((InStr(34, strLinha, vbTab) - 1) - 34)))
                D6 = Trim(Mid(strLinha, 38, ((InStr(38, strLinha, vbTab) - 1) - 38)))
            End If
            
            Result.AddNew
            Result("N�mero") = N�mero
            Result("1D") = D1
            Result("2D") = D2
            Result("3D") = D3
            Result("4D") = D4
            Result("5D") = D5
            Result("6D") = D6
            Result.Update
        Wend
    End If
End Sub

Private Sub Command2_Click()
    Deleta = "Delete * from Tempor�rio"
    Banco.Execute Deleta
    
    Tempor�rio.Index = "Chave1"
    Result.Index = "Chave1"
    Result.MoveFirst
    While Not Result.EOF
        Tempor�rio.Seek "=", Result("1D")
        If Tempor�rio.NoMatch Then
            Tempor�rio.AddNew
            Tempor�rio("N�mero") = Result("1D")
            Tempor�rio("Vezes") = 0
        Else
            Tempor�rio.Edit
        End If
        Tempor�rio("Vezes") = Tempor�rio("Vezes") + 1
        Tempor�rio.Update
            
        Tempor�rio.Seek "=", Result("2D")
        If Tempor�rio.NoMatch Then
            Tempor�rio.AddNew
            Tempor�rio("N�mero") = Result("2D")
            Tempor�rio("Vezes") = 0
        Else
            Tempor�rio.Edit
        End If
        Tempor�rio("Vezes") = Tempor�rio("Vezes") + 1
        Tempor�rio.Update
            
        Tempor�rio.Seek "=", Result("3D")
        If Tempor�rio.NoMatch Then
            Tempor�rio.AddNew
            Tempor�rio("N�mero") = Result("3D")
            Tempor�rio("Vezes") = 0
        Else
            Tempor�rio.Edit
        End If
        Tempor�rio("Vezes") = Tempor�rio("Vezes") + 1
        Tempor�rio.Update
            
        Tempor�rio.Seek "=", Result("4D")
        If Tempor�rio.NoMatch Then
            Tempor�rio.AddNew
            Tempor�rio("N�mero") = Result("4D")
            Tempor�rio("Vezes") = 0
        Else
            Tempor�rio.Edit
        End If
        Tempor�rio("Vezes") = Tempor�rio("Vezes") + 1
        Tempor�rio.Update
        
        Tempor�rio.Seek "=", Result("5D")
        If Tempor�rio.NoMatch Then
            Tempor�rio.AddNew
            Tempor�rio("N�mero") = Result("5D")
            Tempor�rio("Vezes") = 0
        Else
            Tempor�rio.Edit
        End If
        Tempor�rio("Vezes") = Tempor�rio("Vezes") + 1
        Tempor�rio.Update
        
        Tempor�rio.Seek "=", Result("6D")
        If Tempor�rio.NoMatch Then
            Tempor�rio.AddNew
            Tempor�rio("N�mero") = Result("6D")
            Tempor�rio("Vezes") = 0
        Else
            Tempor�rio.Edit
        End If
        Tempor�rio("Vezes") = Tempor�rio("Vezes") + 1
        Tempor�rio.Update
        
        Result.MoveNext
    Wend
    
    
    MSFlexGrid1.Clear
    MSFlexGrid1.FormatString = "|^N�mero|^Vezes sorteadas "
    MSFlexGrid1.Rows = 1
    Tempor�rio.Index = "Chave2"
    Tempor�rio.MoveLast
    While Not Tempor�rio.BOF
        MSFlexGrid1.AddItem Chr(9) & Tempor�rio("N�mero") & Chr(9) & Tempor�rio("Vezes")
        Tempor�rio.MovePrevious
    Wend
End Sub

Private Sub Form_Load()
    Set Banco = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\Banco.mdb")
    Set Result = Banco.OpenRecordset("Resultados", dbOpenTable)
    Set Tempor�rio = Banco.OpenRecordset("Tempor�rio", dbOpenTable)
End Sub
