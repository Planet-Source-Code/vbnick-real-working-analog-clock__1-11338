VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   450
      Top             =   690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Done By BenchMarker Kayhan Tanriseven....If you like the code please vote...the improved version is coming out soon!"
      Height          =   870
      Left            =   900
      TabIndex        =   0
      Top             =   3135
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_load()
Width = ScaleHeight
Scale (-20, 20)-(20, -20)
Timer1.Interval = 1000 'equals to 1 second
AutoRedraw = True
Circle (0, 0), 20
DrawMode = 7 'xor
End Sub
Private Sub timer1_timer()
Dim degrees, seconds, minutes, hours
Static sx, sy, dx, dy, stx, sty
Caption = Time
DrawWidth = 1
Line (0, 0)-(sx, sy), QBColor(10) 'go over the last one by deleting it and drawing new hand
seconds = second(Time) 'take seconds from the timer
degrees = -seconds * 6 + 90 'every one second is 6 degrees. 360degree/60second
sx = 20 * Cos(degrees * 3.1415 / 180)
sy = 20 * Sin(degrees * 3.1415 / 180)
Line (0, 0)-(sx, sy), QBColor(10) 'draw the second hand
DrawWidth = 3
Line (0, 0)-(dx, dy), QBColor(11)
minutes = minute(Time)
degrees = -minutes * 6 + 90 'every one minute is 6 degrees. 360degree/60second
dx = 20 * Cos(degrees * 3.1415 / 180)
dy = 20 * Sin(degrees * 3.1415 / 180)
Line (0, 0)-(dx, dy), QBColor(11) 'draw minutes hand
DrawWidth = 3
Line (0, 0)-(stx, sty), QBColor(12)
hours = hour(Time)
degrees = -hours * 30 + 90 'every one hour is 30 degrees. 360degree/12hours = 30
stx = 12 * Cos(degrees * 3.1415 / 180)
sty = 12 * Sin(degrees * 3.1415 / 180)
Line (0, 0)-(stx, sty), QBColor(12) 'draw hour hand
End Sub
