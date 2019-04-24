VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   573.63
   ScaleMode       =   0  'User
   ScaleWidth      =   420
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   4320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "INSTRUCTIONS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
      Width           =   4335
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "IRONMAN SPEED +2 At every 100'th SCORE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Catch FUEL          Score=6"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Catch ALIENS      Score=-3"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Move RIGHT D    Move LEFT  S"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   3600
      Picture         =   "Ironman.frx":0000
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   4710
      Picture         =   "Ironman.frx":22CB1
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   1950
      Picture         =   "Ironman.frx":6033B
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   3187
      Picture         =   "Ironman.frx":9D9C5
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SPEED"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IRONMAN"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2670
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, c, score, speed As Integer
Dim a, f As String
Dim enms(2), food(2) As Image



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyD Then
Image4.Left = Image4.Left + speed
ElseIf KeyCode = vbKeyA Then
Image4.Left = Image4.Left - speed
End If
End Sub

Private Sub Form_Load()
c = 2
speed = 5
score = 0
a = "enemies"
f = "foods"
End Sub

Private Sub Image1_Click()
Label1.Visible = True
Label2.Visible = True
Label8.Visible = True
Label9.Visible = True
Label3.Visible = False
Image1.Visible = False
Image2.Visible = True
Image2.Visible = False
Image3.Visible = False
Frame1.Visible = False
createenemies (2)
createfood (2)
Timer1.Enabled = True
End Sub

 
Private Sub Timer1_Timer()
Label2.Caption = score
Label9.Caption = speed
For i = 0 To c
enms(i).Visible = True
food(i).Visible = True
enms(i).Top = enms(i).Top + 30
food(i).Top = food(i).Top + 30
If collision(enms(i), Image4) = True Then
score = score - 1
enms(i).Visible = False
End If
If collision(food(i), Image4) = True Then
food(i).Visible = False
score = score + 2
If score Mod 100 = 0 Then
speed = speed + 1
End If
End If
If enms(i).Top >= 600 Then
Randomize
enms(i).Left = (450 * Rnd())
food(i).Left = (450 * Rnd())
enms(i).Top = enms(i).Top - enms(i).Top + 1
food(i).Top = food(i).Top - food(i).Top + 1
End If
Next
End Sub

Private Function collision(ByVal object1 As Object, ByVal object2 As Object) As Boolean
If object1.Top + object1.Height >= object2.Top And _
   object2.Top + object2.Height >= object1.Top And _
   object1.Left + object1.Width >= object2.Left And _
   object2.Left + object2.Width >= object1.Left And object1.Visible = True And object2.Visible = True Then
   collision = True
   Else
   collision = False
    End If
End Function
Private Sub createenemies(number)
Dim i As Integer
For i = 0 To number
Dim temp As Image
Set temp = Me.Controls.Add("VB.Image", "a" & i)
temp.Top = 2 * Rnd()
temp.Height = 40
temp.Width = 40
Randomize
temp.Left = (450 * Rnd())
temp.Stretch = True
temp.Picture = LoadPicture("C:\Users\admin\Desktop\frog.gif")
temp.Width = 40
temp.Width = 40
Set enms(i) = temp
enms(i).Visible = True
Next
End Sub
Private Sub createfood(number)
Dim i As Integer
For i = 0 To number
Dim temp As Image
Set temp = Me.Controls.Add("VB.Image", "f" & i)
temp.Top = 2 * Rnd()
Randomize
temp.Left = (450 * Rnd())
temp.Stretch = True
temp.Picture = LoadPicture("C:\Users\admin\Desktop\fire.gif")
temp.Width = 40
temp.Width = 40
Set food(i) = temp
food(i).Visible = True
Next
End Sub
