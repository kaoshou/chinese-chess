VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����H��"
   ClientHeight    =   8055
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "���s���"
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���s���"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line30 
      X1              =   4080
      X2              =   5280
      Y1              =   5400
      Y2              =   6600
   End
   Begin VB.Line Line29 
      X1              =   5280
      X2              =   4080
      Y1              =   5400
      Y2              =   6600
   End
   Begin VB.Line Line28 
      X1              =   4080
      X2              =   5280
      Y1              =   1200
      Y2              =   2400
   End
   Begin VB.Line Line27 
      X1              =   4080
      X2              =   5280
      Y1              =   2400
      Y2              =   1200
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   89
      Left            =   6840
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   88
      Left            =   6240
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   87
      Left            =   5640
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   86
      Left            =   5040
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   85
      Left            =   4440
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   84
      Left            =   3840
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   83
      Left            =   3240
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   82
      Left            =   2640
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   81
      Left            =   2040
      Top             =   6360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   80
      Left            =   6840
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   79
      Left            =   6240
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   78
      Left            =   5640
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   77
      Left            =   5040
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   76
      Left            =   4440
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   75
      Left            =   3840
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   74
      Left            =   3240
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   73
      Left            =   2640
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   72
      Left            =   2040
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   71
      Left            =   6840
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   70
      Left            =   6240
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   69
      Left            =   5640
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   68
      Left            =   5040
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   67
      Left            =   4440
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   66
      Left            =   3840
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   65
      Left            =   3240
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   64
      Left            =   2640
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   63
      Left            =   2040
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   62
      Left            =   6840
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   61
      Left            =   6240
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   60
      Left            =   5640
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   59
      Left            =   5040
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   58
      Left            =   4440
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   57
      Left            =   3840
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   56
      Left            =   3240
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   55
      Left            =   2640
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   54
      Left            =   2040
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   53
      Left            =   6840
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   52
      Left            =   6240
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   51
      Left            =   5640
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   50
      Left            =   5040
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   49
      Left            =   4440
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   48
      Left            =   3840
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   47
      Left            =   3240
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   46
      Left            =   2640
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   45
      Left            =   2040
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   44
      Left            =   6840
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   43
      Left            =   6240
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   42
      Left            =   5640
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   41
      Left            =   5040
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   40
      Left            =   4440
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   39
      Left            =   3840
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   38
      Left            =   3240
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   37
      Left            =   2640
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   36
      Left            =   2040
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   35
      Left            =   6840
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   34
      Left            =   6240
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   33
      Left            =   5640
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   32
      Left            =   5040
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   31
      Left            =   4440
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   30
      Left            =   3840
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   29
      Left            =   3240
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   28
      Left            =   2640
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   27
      Left            =   2040
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   26
      Left            =   6840
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   25
      Left            =   6240
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   24
      Left            =   5640
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   23
      Left            =   5040
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   22
      Left            =   4440
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   21
      Left            =   3840
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   20
      Left            =   3240
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   19
      Left            =   2640
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   18
      Left            =   2040
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   17
      Left            =   6840
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   16
      Left            =   6240
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   15
      Left            =   5640
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   14
      Left            =   5040
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   13
      Left            =   4440
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   3840
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   11
      Left            =   3240
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   10
      Left            =   2640
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   9
      Left            =   2040
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   8
      Left            =   6840
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   7
      Left            =   6240
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   5640
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   5040
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   4440
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   3840
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   3240
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   2640
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   2040
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   11
      Left            =   5160
      Picture         =   "Form1.frx":0000
      Tag             =   "�X"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   7680
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Line Line26 
      X1              =   6480
      X2              =   6480
      Y1              =   3600
      Y2              =   1200
   End
   Begin VB.Line Line25 
      X1              =   5280
      X2              =   5280
      Y1              =   3600
      Y2              =   1200
   End
   Begin VB.Line Line24 
      X1              =   5880
      X2              =   5880
      Y1              =   3600
      Y2              =   1200
   End
   Begin VB.Line Line23 
      X1              =   4080
      X2              =   4080
      Y1              =   3600
      Y2              =   1200
   End
   Begin VB.Line Line22 
      X1              =   4680
      X2              =   4680
      Y1              =   3600
      Y2              =   1200
   End
   Begin VB.Line Line21 
      X1              =   7080
      X2              =   2280
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line20 
      X1              =   7080
      X2              =   2280
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line19 
      X1              =   2280
      X2              =   7080
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line18 
      X1              =   7080
      X2              =   2280
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line17 
      X1              =   2880
      X2              =   2880
      Y1              =   3600
      Y2              =   1200
   End
   Begin VB.Line Line16 
      X1              =   3480
      X2              =   3480
      Y1              =   3600
      Y2              =   1200
   End
   Begin VB.Line Line15 
      X1              =   7080
      X2              =   2280
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line14 
      X1              =   6480
      X2              =   6480
      Y1              =   6600
      Y2              =   4200
   End
   Begin VB.Line Line13 
      X1              =   5280
      X2              =   5280
      Y1              =   6600
      Y2              =   4200
   End
   Begin VB.Line Line12 
      X1              =   5880
      X2              =   5880
      Y1              =   6600
      Y2              =   4200
   End
   Begin VB.Line Line11 
      X1              =   4080
      X2              =   4080
      Y1              =   6600
      Y2              =   4200
   End
   Begin VB.Line Line10 
      X1              =   4680
      X2              =   4680
      Y1              =   6600
      Y2              =   4200
   End
   Begin VB.Line Line9 
      X1              =   7080
      X2              =   2280
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line8 
      X1              =   7080
      X2              =   2280
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line7 
      X1              =   7080
      X2              =   2280
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line6 
      X1              =   3480
      X2              =   3480
      Y1              =   6600
      Y2              =   4200
   End
   Begin VB.Line Line5 
      X1              =   7080
      X2              =   7080
      Y1              =   6600
      Y2              =   1200
   End
   Begin VB.Line Line4 
      X1              =   2880
      X2              =   2880
      Y1              =   6600
      Y2              =   4200
   End
   Begin VB.Line Line3 
      X1              =   7080
      X2              =   2280
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   2280
      Y1              =   6600
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   7080
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   13
      Left            =   6120
      Picture         =   "Form1.frx":0C26
      Tag             =   "�L"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   12
      Left            =   5640
      Picture         =   "Form1.frx":1868
      Tag             =   "��"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   4680
      Picture         =   "Form1.frx":248E
      Tag             =   "r��"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   4200
      Picture         =   "Form1.frx":30B4
      Tag             =   "��"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   3720
      Picture         =   "Form1.frx":3CDA
      Tag             =   "�K"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   3240
      Picture         =   "Form1.frx":4900
      Tag             =   "��"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   5880
      Picture         =   "Form1.frx":5526
      Tag             =   "��"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   5400
      Picture         =   "Form1.frx":6168
      Tag             =   "�]"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   4440
      Picture         =   "Form1.frx":6D8E
      Tag             =   "��"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   4920
      Picture         =   "Form1.frx":79B4
      Tag             =   "��"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   3960
      Picture         =   "Form1.frx":85DA
      Tag             =   "�H"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   3480
      Picture         =   "Form1.frx":921C
      Tag             =   "�h"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   3000
      Picture         =   "Form1.frx":9E42
      Tag             =   "�N"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu stgo 
      Caption         =   "�}�l"
      Begin VB.Menu rego 
         Caption         =   "���s�}�l"
      End
      Begin VB.Menu as 
         Caption         =   "����"
      End
      Begin VB.Menu endgame 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'
'    Copyright (C) 2007-2011  Yu-Han Cheng (�G����)
'
'    This program (����H��,chinese-chess) is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'    My E-mail Address is "kaoshou@gmail.com" (Yu-Han Cheng)
'
'=============================================================================

Dim xandy(9, 10) As String '�ѽL
Dim ax, ay As Integer '�O����ܰ��Ѥl
Dim anum As Integer
Dim mace As String '�O�������
Dim okgg As String '�O���N�Ӧ��L�ӭ�


Private Sub Command2_Click()
ax = 0
ay = 0
aum = 0
Command2.Visible = False
La
End Sub
Private Sub Command3_Click()
ax = 0
ay = 0
aum = 0
Command3.Visible = False
La
End Sub

Private Sub endgame_Click()
End
End Sub

Private Sub Form_Load()
afresh
mace = "�¤�"
La
End Sub

Private Sub Image2_Click(Index As Integer)
'�o�q�{�����դ�
If randb(Index) = mace And ax = 0 And ay = 0 Then
    ax = numtox(Index)
    ay = numtoy(Index)
    anum = Index
ElseIf ax <> 0 And ay <> 0 And (ax <> numtox(Index) Or ay <> numtoy(Index)) Then
    Select Case xandy(ax, ay)
    Case "�N", "��"
    Go1 (Index)
    Case "�h", "�K"
    Go2 (Index)
    Case "�H", "��"
    Go3 (Index)
    Case "��", "�X"
    Go4 (Index)
    Case "��", "r��"
    Go5 (Index)
    Case "�]", "��"
    Go6 (Index)
    Case "��", "�L"
    Go7 (Index)
    End Select
End If
La
Ending
End Sub

Sub Go1(num As Integer)
'�N�Ӥ�����
If numtox(num) = 4 Or numtox(num) = 5 Or numtox(num) = 6 Then
If (ax = numtox(num) And (ay - numtoy(num) = 1 Or ay - numtoy(num) = -1)) Or (ay = numtoy(num) And (ax - numtox(num) = 1 Or ax - numtox(num) = -1)) Then
    If mace = "�¤�" Then
        If numtoy(num) = 1 Or numtoy(num) = 2 Or numtoy(num) = 3 Then
                If randb(num) = "����" Or randb(num) = "" Then
                    emend (num)
                End If
        End If
    ElseIf mace = "����" Then
        If numtoy(num) = 8 Or numtoy(num) = 9 Or numtoy(num) = 10 Then
                    If randb(num) = "�¤�" Or randb(num) = "" Then
                    emend (num)
                End If
        End If
    End If
End If
End If
End Sub
Sub Go2(num As Integer)
'�h�K������
If numtox(num) = 4 Or numtox(num) = 5 Or numtox(num) = 6 Then
If ax <> numtox(num) And ay <> numtoy(num) Then
    If mace = "�¤�" Then
        If numtoy(num) = 1 Or numtoy(num) = 2 Or numtoy(num) = 3 Then
            If ((ax + 1 = numtox(num)) Or (ax - 1 = numtox(num))) And ((ay + 1 = numtoy(num)) Or (ay - 1 = numtoy(num))) Then
                If randb(num) = "����" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        End If
    ElseIf mace = "����" Then
        If numtoy(num) = 8 Or numtoy(num) = 9 Or numtoy(num) = 10 Then
            If ((ax + 1 = numtox(num)) Or (ax - 1 = numtox(num))) And ((ay + 1 = numtoy(num)) Or (ay - 1 = numtoy(num))) Then
                    If randb(num) = "�¤�" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        End If
    End If
End If
End If
End Sub
Sub Go3(num As Integer)
'�H�ۤ�����
If ax <> numtox(num) Or ay <> numtoy(num) Then
    If mace = "�¤�" Then
        If numtoy(num) <= 5 Then
            If ((ax + 2 = numtox(num)) Or (ax - 2 = numtox(num))) And ((ay + 2 = numtoy(num)) Or (ay - 2 = numtoy(num))) Then
                If (randb(num) = "����" Or randb(num) = "") And (xandy((ax + numtox(num)) / 2, (ay + numtoy(num)) / 2) = "") Then '�B�Τ��I�����P�_���L��в�
                    emend (num)
                End If
            End If
        End If
    ElseIf mace = "����" Then
        If numtoy(num) >= 6 Then
            If ((ax + 2 = numtox(num)) Or (ax - 2 = numtox(num))) And ((ay + 2 = numtoy(num)) Or (ay - 2 = numtoy(num))) Then
                    If (randb(num) = "�¤�" Or randb(num) = "") And (xandy((ax + numtox(num)) / 2, (ay + numtoy(num)) / 2) = "") Then
                    emend (num)
                End If
            End If
        End If
    End If
End If
End Sub
Sub Go4(num As Integer)
'���X������
If ax <> numtox(num) Or ay <> numtoy(num) Then
    If (ax + 2 = numtox(num) Or ax - 2 = numtox(num)) And (ay + 1 = numtoy(num) Or ay - 1 = numtoy(num)) Then
        If xandy((ax + numtox(num)) / 2, ay) = "" Then
            If randb(num) <> mace Or randb(num) = "" Then
                    emend (num)
            End If
        End If
    ElseIf (ax + 1 = numtox(num) Or ax - 1 = numtox(num)) And (ay + 2 = numtoy(num) Or ay - 2 = numtoy(num)) Then
        If xandy(ax, (ay + numtoy(num)) / 2) = "" Then
            If randb(num) <> mace Or randb(num) = "" Then
                    emend (num)
            End If
        End If
    End If
End If
End Sub
Sub Go5(num As Integer)
'��r��������
If ax <> numtox(num) Or ay <> numtoy(num) Then
    If ax = numtox(num) Then
        If ay < numtoy(num) Then
            i = 1
        Else
            i = -1
        End If
        
        
        For fory = ay + i To numtoy(num) Step i
            If xandy(ax, fory) = "" Or (fory = numtoy(num)) Then
                If fory = numtoy(num) And (randb(num) <> mace Or randb(num) = "") Then
                    emend (num)
                 End If
            Else
                Exit For
            End If
        Next fory
        
        
        
    ElseIf ay = numtoy(num) Then
        If ax < numtox(num) Then
            i = 1
        Else
            i = -1
        End If
        For forx = ax + i To numtox(num) Step i
            If xandy(forx, ay) = "" Or (forx = numtox(num)) Then
                If forx = numtox(num) And (randb(num) <> mace Or randb(num) = "") Then
                    emend (num)
                End If
            Else
            Exit For
            End If
        Next forx
    End If
End If
End Sub

Sub Go6(num As Integer)
'�]��������
If ax <> numtox(num) Or ay <> numtoy(num) Then
If xandy(numtox(num), numtoy(num)) = "" Then
'��������
    If ax = numtox(num) Then
        If ay < numtoy(num) Then
            i = 1
        Else
            i = -1
        End If
        
        
        For fory = ay + i To numtoy(num) Step i
            If xandy(ax, fory) = "" Then
                If fory = numtoy(num) Then
                    emend (num)
                 End If
            Else
                Exit For
            End If
        Next fory
        
        
        
    ElseIf ay = numtoy(num) Then
        If ax < numtox(num) Then
            i = 1
        Else
            i = -1
        End If
        For forx = ax + i To numtox(num) Step i
            If xandy(forx, ay) = "" Then
                If forx = numtox(num) Then
                    emend (num)
                End If
            Else
            Exit For
            End If
        Next forx
    End If
    
ElseIf xandy(numtox(num), numtoy(num)) <> "" And randb(num) <> mace Then
'�Y������
    a = 0
    If ax = numtox(num) Then
        If ay < numtoy(num) Then
            i = 1
        Else
            i = -1
        End If
        
        
        For fory = ay + i To numtoy(num) Step i
            If xandy(ax, fory) = "" Or fory = numtoy(num) Then
                If fory = numtoy(num) And a = 1 And (randb(num) <> mace) Then
                    emend (num)
                 End If
            Else
                a = a + 1
                If a >= 2 Then
                    Exit For
                End If
            End If
        Next fory
    ElseIf ay = numtoy(num) Then
        If ax < numtox(num) Then
            i = 1
        Else
            i = -1
        End If
        For forx = ax + i To numtox(num) Step i
            If xandy(forx, ay) = "" Or forx = numtox(num) Then
                If forx = numtox(num) And a = 1 And (randb(num) <> mace) Then
                    emend (num)
                 End If
            Else
                a = a + 1
                If a >= 2 Then
                    Exit For
                End If
            End If
        Next forx
    End If

End If
End If
End Sub
Sub Go7(num As Integer)
'��L������
If ax <> numtox(num) Or ay <> numtoy(num) Then
    If mace = "�¤�" Then
        If ay < 6 Then
            If numtoy(num) = ay + 1 And numtox(num) = ax Then
                If randb(num) = "����" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        Else
            If (numtoy(num) = ay + 1 And numtox(num) = ax) Or (numtox(num) = ax + 1 And numtoy(num) = ay) Or (numtox(num) = ax - 1 And numtoy(num) = ay) Then
                If randb(num) = "����" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        End If
    ElseIf mace = "����" Then
        If ay > 5 Then
            If numtoy(num) = ay - 1 And numtox(num) = ax Then
                If randb(num) = "�¤�" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        Else
            If (numtoy(num) = ay - 1 And numtox(num) = ax) Or (numtox(num) = ax + 1 And numtoy(num) = ay) Or (numtox(num) = ax - 1 And numtoy(num) = ay) Then
                If randb(num) = "�¤�" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        End If
    End If
End If
End Sub
Sub start()
'���y�ѽL�A���s�t�m
a = 0
For y = 1 To 10 Step 1
    For x = 1 To 9 Step 1
        If xandy(x, y) <> "" Then
            For i = 0 To 13 Step 1
                If Image1(i).Tag = xandy(x, y) Then
                    Image2(a).Picture = Image1(i).Picture
                    Exit For
                End If
            Next i
        Else
            Image2(a).Picture = LoadPicture
        End If
        a = a + 1
    Next x
Next
End Sub
Sub Ending()
'�˴��ӭt
    For a = 1 To 9 Step 1
        For b = 1 To 10 Step 1
            If xandy(a, b) = "�N" Then
                Blackend = 1
            End If
            If xandy(a, b) = "��" Then
                Rndend = 1
            End If
        Next b
    Next a
If Blackend = 0 Or okgg = "�¤�" Then
    over = MsgBox("����ӡA�¤��", , "�C������")
    Label1.Caption = "�¤��"
    Label2.Caption = "�����"
    Command2.Visible = False
    Command3.Visible = False
End If
If Rndend = 0 Or okgg = "����" Then
    over = MsgBox("�¤�ӡA�����", , "�C������")
    Label1.Caption = "�¤��"
    Label2.Caption = "�����"
    Command2.Visible = False
    Command3.Visible = False
End If
End Sub
Public Function numtox(num As Integer)
'0~89�}�C�Ʀr �� X,Y�G���}�C ���oX�� (num���I�蠟0-89�}�C)
numtox = (num Mod 9) + 1
End Function
Public Function numtoy(num As Integer)
'0~89�}�C�Ʀr �� X,Y�G���}�C ���oY�� (num���I�蠟0-89�}�C)
numtoy = 1 + (num \ 9)
End Function
Public Function iswhat(num As Integer)
'0~89�}�C�Ʀr �� �I��Ѥl�W�� (num���I�蠟0-89�}�C)
iswhat = xandy(numtox(num), numtoy(num))
End Function
Public Function randb(num As Integer)
'0~89�}�C�Ʀr �P�_�O�¬O�� (num���I�蠟0-89�}�C)
If xandy(numtox(num), numtoy(num)) <> "" Then
    For i = 0 To 13 Step 1
        If Image1(i).Tag = xandy(numtox(num), numtoy(num)) Then
            If i < 7 Then
                randb = "�¤�"
            Else
                randb = "����"
            End If
            Exit For
        End If
    Next i
End If
End Function
Sub emend(num As Integer)
'��Ѥl�M�w���ʨ�Y��m�]���ΦY�^��A�Φ���Ʋ��ʴѤl�A�çⱱ���v�浹�t�@��C
'���Ƶ{���t��start�Ƶ{���A�G�|���s��z�ѽL�C
xandy(numtox(num), numtoy(num)) = xandy(ax, ay)
xandy(ax, ay) = ""
start
ax = 0: ay = 0: anum = 0
If okg = 1 Then
okgg = mace
End If
If mace = "�¤�" Then
    mace = "����"
Else
    mace = "�¤�"
End If
Command2.Visible = False
Command3.Visible = False
End Sub
Sub La()
'�������Label
If mace = "�¤�" Then
    Label2.Caption = ""
    Label1.Caption = "����¤�" & Chr(13) + Chr(10)
    If ax <> 0 And ay <> 0 Then
        Label1.Caption = Label1.Caption & "�@�w��ܤF " & xandy(ax, ay)
        Command2.Visible = True
    End If
ElseIf mace = "����" Then
    Label1.Caption = ""
    Label2.Caption = "�������" & Chr(13) + Chr(10)
    If ax <> 0 And ay <> 0 Then
        If xandy(ax, ay) <> "r��" Then
            Label2.Caption = Label2.Caption & "�@�w��ܤF " & xandy(ax, ay)
        ElseIf xandy(ax, ay) = "r��" Then
            Label2.Caption = Label2.Caption & "�@�w��ܤF ��"
        End If
        Command3.Visible = True
    End If
End If
End Sub
Public Function okg()
'�˴��N�Ӧ��L�b�P�@���u
    okg = 0
    For a = 1 To 9 Step 1
        For b = 1 To 10 Step 1
            If xandy(a, b) = "�N" Then
                nx = a
                ny = b
            End If
            If xandy(a, b) = "��" Then
                mx = a
                my = b
            End If
        Next b
    Next a
    If nx = mx Then
    For forxy = ny + 1 To my - 1 Step 1
        If xandy(nx, forxy) <> "" Then
            Exit For
        End If
        If xandy(nx, forxy) = "" And forxy = my - 1 Then
            okg = 1
        End If
    Next forxy
    End If
End Function
Sub afresh()
'��l�ƥ�
Label1.Caption = ""
Label2.Caption = ""
okgg = ""
ax = 0: ay = 0: anum = 0: mace = "�¤�"
Command2.Visible = False
Command3.Visible = False
    For a = 1 To 9 Step 1
        For b = 1 To 10 Step 1
            xandy(a, b) = ""
        Next b
    Next a
xandy(1, 1) = "��"
xandy(2, 1) = "��"
xandy(3, 1) = "�H"
xandy(4, 1) = "�h"
xandy(5, 1) = "�N"
xandy(6, 1) = "�h"
xandy(7, 1) = "�H"
xandy(8, 1) = "��"
xandy(9, 1) = "��"
xandy(2, 3) = "�]"
xandy(8, 3) = "�]"
xandy(1, 4) = "��"
xandy(3, 4) = "��"
xandy(5, 4) = "��"
xandy(7, 4) = "��"
xandy(9, 4) = "��"
xandy(1, 10) = "r��"
xandy(2, 10) = "�X"
xandy(3, 10) = "��"
xandy(4, 10) = "�K"
xandy(5, 10) = "��"
xandy(6, 10) = "�K"
xandy(7, 10) = "��"
xandy(8, 10) = "�X"
xandy(9, 10) = "r��"
xandy(2, 8) = "��"
xandy(8, 8) = "��"
xandy(1, 7) = "�L"
xandy(3, 7) = "�L"
xandy(5, 7) = "�L"
xandy(7, 7) = "�L"
xandy(9, 7) = "�L"
start
End Sub
Private Sub as_Click()
MsgBox ("chinese-chess ����H��  Copyright (C) 2007-2011  Yu-Han Cheng �G����" & vbCrLf & _
        "This program comes with ABSOLUTELY NO WARRANTY" & vbCrLf & _
        "This is free software, and you are welcome to redistribute it" & vbCrLf & _
        "GPLv3 see <http://www.gnu.org/licenses/>")

End Sub
Private Sub rego_Click()
afresh
La
End Sub
