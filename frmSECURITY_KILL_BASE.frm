VERSION 5.00

'		***************************************************************************
'       *    					DATABASE ARMAGEDON DESTRUCTOR				      *
'       * 		hACKING frame developped on VB6.0 by :							  *
'		*																		  *
'		*							Marquos LEDAU RIETTE						  *
'		*																		  *
'		*																		  *
'       * 								2016 - 03								  *
'       ***************************************************************************

Option Explicit


Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSECURITY_KILL_BASE 
   BackColor       =   &H80000014&
   Caption         =   "Générateur de code"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3240
   Icon            =   "frmSECURITY_KILL_BASE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3240
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Portail de sécurité :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000011&
         Caption         =   "Detruire la base de données"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtCODE 
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Code de déstruction de la base :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   2880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ce code vous seras demandé si vous voullez détruire la base de données : BaseDesVins.mdb"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblCODE 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmSECURITY_KILL_BASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
' Declare la variable Var1
Dim Var1 As String
Dim Var2 As String
'Ouvre le fichier
Open "C:\WINDOWS\SYSTEM\Bacchus2-4_KILLDatBas.dll" For Input As #1
'Lit la première ligne et la place dans Var1
Input #1, Var1
Input #1, Var2
'Ferme le fichier
Close #1

'Test si le code est valide
If Var2 = txtCODE.Text Then 'le pass du fichier doit etre = au mot de passe entré dans le text3
    
    Unload Me

    rep = MsgBox("La base de données : BaseDesVins.mdb va être détruite, voullez-vous continuer?", vbYesNo, "Bacchus V" & Version)
    If rep = 6 Then
        On Error Resume Next
        
        
        '1)On test la valeur de l'option de sauvegarde: TRUE ou FALSE
        ' Declare la variable Var1
        'Dim Var3 As String
        'Dim Var2 As String
        'Ouvre le fichier
        Open "C:\WINDOWS\SYSTEM\Bacchus2-4_CopyBas.dll" For Input As #1
        'Lit la première ligne et la place dans Var1
        Input #1, Var1
        Input #1, Var2
        'Ferme le fichier
        Close #1
            If Var2 = "TRUE" Then
            'on copie le fichier
                Dim f As New Scripting.FileSystemObject
                f.CopyFile App.Path & "\BaseDesVins.mdb", App.Path & "\Mes_Archives\BaseDesVins.mdb", True
            ElseIf Var2 = "FALSE" Then
            End If
        
        '2)On détruit la base et le fichier contenant le code de destruction
        Kill (App.Path & "\BaseDesVins.mdb")
        Kill ("C:\WINDOWS\SYSTEM\Bacchus2-4_KILLDatBas.dll")
        Unload MDIForm1 'on decharge la MDI
        MsgBox "La base de données est détruite." & vbNewLine & "Bacchus V " & Version & " ne pourra pas fonctionner correctement, le programme va donc s'arrêter." & vbNewLine & "Pour refaire fonctionner le programme correctement, relancez l'application et re-créez une nouvelle base de données: BaseDesVins.mdb", vbInformation
        End 'Si la base est détruite, on arrête le programme
    
    ElseIf rep = 7 Then
        Exit Sub
    End If
    
ElseIf Var2 <> txtCODE.Text Then 'fin si
    txtCODE.Text = ""  'si pass mauvais on efface text2
    Label2.Caption = "Error:" & vbNewLine & "Mot de passe invalide"  'et on ecrit mauvais mot de passe
    txtCODE.SetFocus
End If

End Sub


Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Timer1.Interval = 300
    ProgressBar1.Value = 0
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + Val(1)

If ProgressBar1.Value = 100 Then
Timer1.Enabled = False

Unload Me
End If

End Sub
