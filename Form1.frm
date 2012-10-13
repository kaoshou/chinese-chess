VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "中國象棋"
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
      Caption         =   "重新選擇"
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新選擇"
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
      Tag             =   "傌"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
      Tag             =   "兵"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   12
      Left            =   5640
      Picture         =   "Form1.frx":1868
      Tag             =   "炮"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   4680
      Picture         =   "Form1.frx":248E
      Tag             =   "r車"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   4200
      Picture         =   "Form1.frx":30B4
      Tag             =   "相"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   3720
      Picture         =   "Form1.frx":3CDA
      Tag             =   "仕"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   3240
      Picture         =   "Form1.frx":4900
      Tag             =   "帥"
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   5880
      Picture         =   "Form1.frx":5526
      Tag             =   "卒"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   5400
      Picture         =   "Form1.frx":6168
      Tag             =   "包"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   4440
      Picture         =   "Form1.frx":6D8E
      Tag             =   "車"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   4920
      Picture         =   "Form1.frx":79B4
      Tag             =   "馬"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   3960
      Picture         =   "Form1.frx":85DA
      Tag             =   "象"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   3480
      Picture         =   "Form1.frx":921C
      Tag             =   "士"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   3000
      Picture         =   "Form1.frx":9E42
      Tag             =   "將"
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu stgo 
      Caption         =   "開始"
      Begin VB.Menu rego 
         Caption         =   "重新開始"
      End
      Begin VB.Menu as 
         Caption         =   "關於"
      End
      Begin VB.Menu endgame 
         Caption         =   "結束"
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
'    Copyright (C) 2007-2011  Yu-Han Cheng (鄭郁翰)
'
'    This program (中國象棋,chinese-chess) is free software: you can redistribute it and/or modify
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

Dim xandy(9, 10) As String '棋盤
Dim ax, ay As Integer '記錄選擇隻棋子
Dim anum As Integer
Dim mace As String '記錄輪到誰
Dim okgg As String '記錄將帥有無照面


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
mace = "黑方"
La
End Sub

Private Sub Image2_Click(Index As Integer)
'這段程式測試中
If randb(Index) = mace And ax = 0 And ay = 0 Then
    ax = numtox(Index)
    ay = numtoy(Index)
    anum = Index
ElseIf ax <> 0 And ay <> 0 And (ax <> numtox(Index) Or ay <> numtoy(Index)) Then
    Select Case xandy(ax, ay)
    Case "將", "帥"
    Go1 (Index)
    Case "士", "仕"
    Go2 (Index)
    Case "象", "相"
    Go3 (Index)
    Case "馬", "傌"
    Go4 (Index)
    Case "車", "r車"
    Go5 (Index)
    Case "包", "炮"
    Go6 (Index)
    Case "卒", "兵"
    Go7 (Index)
    End Select
End If
La
Ending
End Sub

Sub Go1(num As Integer)
'將帥之移動
If numtox(num) = 4 Or numtox(num) = 5 Or numtox(num) = 6 Then
If (ax = numtox(num) And (ay - numtoy(num) = 1 Or ay - numtoy(num) = -1)) Or (ay = numtoy(num) And (ax - numtox(num) = 1 Or ax - numtox(num) = -1)) Then
    If mace = "黑方" Then
        If numtoy(num) = 1 Or numtoy(num) = 2 Or numtoy(num) = 3 Then
                If randb(num) = "紅方" Or randb(num) = "" Then
                    emend (num)
                End If
        End If
    ElseIf mace = "紅方" Then
        If numtoy(num) = 8 Or numtoy(num) = 9 Or numtoy(num) = 10 Then
                    If randb(num) = "黑方" Or randb(num) = "" Then
                    emend (num)
                End If
        End If
    End If
End If
End If
End Sub
Sub Go2(num As Integer)
'士仕之移動
If numtox(num) = 4 Or numtox(num) = 5 Or numtox(num) = 6 Then
If ax <> numtox(num) And ay <> numtoy(num) Then
    If mace = "黑方" Then
        If numtoy(num) = 1 Or numtoy(num) = 2 Or numtoy(num) = 3 Then
            If ((ax + 1 = numtox(num)) Or (ax - 1 = numtox(num))) And ((ay + 1 = numtoy(num)) Or (ay - 1 = numtoy(num))) Then
                If randb(num) = "紅方" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        End If
    ElseIf mace = "紅方" Then
        If numtoy(num) = 8 Or numtoy(num) = 9 Or numtoy(num) = 10 Then
            If ((ax + 1 = numtox(num)) Or (ax - 1 = numtox(num))) And ((ay + 1 = numtoy(num)) Or (ay - 1 = numtoy(num))) Then
                    If randb(num) = "黑方" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        End If
    End If
End If
End If
End Sub
Sub Go3(num As Integer)
'象相之移動
If ax <> numtox(num) Or ay <> numtoy(num) Then
    If mace = "黑方" Then
        If numtoy(num) <= 5 Then
            If ((ax + 2 = numtox(num)) Or (ax - 2 = numtox(num))) And ((ay + 2 = numtoy(num)) Or (ay - 2 = numtoy(num))) Then
                If (randb(num) = "紅方" Or randb(num) = "") And (xandy((ax + numtox(num)) / 2, (ay + numtoy(num)) / 2) = "") Then '運用中點公式判斷有無塞田眼
                    emend (num)
                End If
            End If
        End If
    ElseIf mace = "紅方" Then
        If numtoy(num) >= 6 Then
            If ((ax + 2 = numtox(num)) Or (ax - 2 = numtox(num))) And ((ay + 2 = numtoy(num)) Or (ay - 2 = numtoy(num))) Then
                    If (randb(num) = "黑方" Or randb(num) = "") And (xandy((ax + numtox(num)) / 2, (ay + numtoy(num)) / 2) = "") Then
                    emend (num)
                End If
            End If
        End If
    End If
End If
End Sub
Sub Go4(num As Integer)
'馬傌之移動
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
'車r車之移動
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
'包炮之移動
If ax <> numtox(num) Or ay <> numtoy(num) Then
If xandy(numtox(num), numtoy(num)) = "" Then
'走的部分
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
'吃的部分
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
'卒兵的移動
If ax <> numtox(num) Or ay <> numtoy(num) Then
    If mace = "黑方" Then
        If ay < 6 Then
            If numtoy(num) = ay + 1 And numtox(num) = ax Then
                If randb(num) = "紅方" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        Else
            If (numtoy(num) = ay + 1 And numtox(num) = ax) Or (numtox(num) = ax + 1 And numtoy(num) = ay) Or (numtox(num) = ax - 1 And numtoy(num) = ay) Then
                If randb(num) = "紅方" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        End If
    ElseIf mace = "紅方" Then
        If ay > 5 Then
            If numtoy(num) = ay - 1 And numtox(num) = ax Then
                If randb(num) = "黑方" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        Else
            If (numtoy(num) = ay - 1 And numtox(num) = ax) Or (numtox(num) = ax + 1 And numtoy(num) = ay) Or (numtox(num) = ax - 1 And numtoy(num) = ay) Then
                If randb(num) = "黑方" Or randb(num) = "" Then
                    emend (num)
                End If
            End If
        End If
    End If
End If
End Sub
Sub start()
'掃描棋盤，重新配置
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
'檢測勝負
    For a = 1 To 9 Step 1
        For b = 1 To 10 Step 1
            If xandy(a, b) = "將" Then
                Blackend = 1
            End If
            If xandy(a, b) = "帥" Then
                Rndend = 1
            End If
        Next b
    Next a
If Blackend = 0 Or okgg = "黑方" Then
    over = MsgBox("紅方勝，黑方敗", , "遊戲結束")
    Label1.Caption = "黑方敗"
    Label2.Caption = "紅方勝"
    Command2.Visible = False
    Command3.Visible = False
End If
If Rndend = 0 Or okgg = "紅方" Then
    over = MsgBox("黑方勝，紅方敗", , "遊戲結束")
    Label1.Caption = "黑方勝"
    Label2.Caption = "紅方敗"
    Command2.Visible = False
    Command3.Visible = False
End If
End Sub
Public Function numtox(num As Integer)
'0~89陣列數字 轉 X,Y二維陣列 取得X值 (num為點選之0-89陣列)
numtox = (num Mod 9) + 1
End Function
Public Function numtoy(num As Integer)
'0~89陣列數字 轉 X,Y二維陣列 取得Y值 (num為點選之0-89陣列)
numtoy = 1 + (num \ 9)
End Function
Public Function iswhat(num As Integer)
'0~89陣列數字 轉 點選棋子名稱 (num為點選之0-89陣列)
iswhat = xandy(numtox(num), numtoy(num))
End Function
Public Function randb(num As Integer)
'0~89陣列數字 判斷是黑是紅 (num為點選之0-89陣列)
If xandy(numtox(num), numtoy(num)) <> "" Then
    For i = 0 To 13 Step 1
        If Image1(i).Tag = xandy(numtox(num), numtoy(num)) Then
            If i < 7 Then
                randb = "黑方"
            Else
                randb = "紅方"
            End If
            Exit For
        End If
    Next i
End If
End Function
Sub emend(num As Integer)
'當棋子決定移動到某位置（走或吃）後，用此函數移動棋子，並把控制權交給另一方。
'此副程式含有start副程式，故會重新整理棋盤。
xandy(numtox(num), numtoy(num)) = xandy(ax, ay)
xandy(ax, ay) = ""
start
ax = 0: ay = 0: anum = 0
If okg = 1 Then
okgg = mace
End If
If mace = "黑方" Then
    mace = "紅方"
Else
    mace = "黑方"
End If
Command2.Visible = False
Command3.Visible = False
End Sub
Sub La()
'控制顯示Label
If mace = "黑方" Then
    Label2.Caption = ""
    Label1.Caption = "輪到黑方" & Chr(13) + Chr(10)
    If ax <> 0 And ay <> 0 Then
        Label1.Caption = Label1.Caption & "　已選擇了 " & xandy(ax, ay)
        Command2.Visible = True
    End If
ElseIf mace = "紅方" Then
    Label1.Caption = ""
    Label2.Caption = "輪到紅方" & Chr(13) + Chr(10)
    If ax <> 0 And ay <> 0 Then
        If xandy(ax, ay) <> "r車" Then
            Label2.Caption = Label2.Caption & "　已選擇了 " & xandy(ax, ay)
        ElseIf xandy(ax, ay) = "r車" Then
            Label2.Caption = Label2.Caption & "　已選擇了 車"
        End If
        Command3.Visible = True
    End If
End If
End Sub
Public Function okg()
'檢測將帥有無在同一直線
    okg = 0
    For a = 1 To 9 Step 1
        For b = 1 To 10 Step 1
            If xandy(a, b) = "將" Then
                nx = a
                ny = b
            End If
            If xandy(a, b) = "帥" Then
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
'初始化用
Label1.Caption = ""
Label2.Caption = ""
okgg = ""
ax = 0: ay = 0: anum = 0: mace = "黑方"
Command2.Visible = False
Command3.Visible = False
    For a = 1 To 9 Step 1
        For b = 1 To 10 Step 1
            xandy(a, b) = ""
        Next b
    Next a
xandy(1, 1) = "車"
xandy(2, 1) = "馬"
xandy(3, 1) = "象"
xandy(4, 1) = "士"
xandy(5, 1) = "將"
xandy(6, 1) = "士"
xandy(7, 1) = "象"
xandy(8, 1) = "馬"
xandy(9, 1) = "車"
xandy(2, 3) = "包"
xandy(8, 3) = "包"
xandy(1, 4) = "卒"
xandy(3, 4) = "卒"
xandy(5, 4) = "卒"
xandy(7, 4) = "卒"
xandy(9, 4) = "卒"
xandy(1, 10) = "r車"
xandy(2, 10) = "傌"
xandy(3, 10) = "相"
xandy(4, 10) = "仕"
xandy(5, 10) = "帥"
xandy(6, 10) = "仕"
xandy(7, 10) = "相"
xandy(8, 10) = "傌"
xandy(9, 10) = "r車"
xandy(2, 8) = "炮"
xandy(8, 8) = "炮"
xandy(1, 7) = "兵"
xandy(3, 7) = "兵"
xandy(5, 7) = "兵"
xandy(7, 7) = "兵"
xandy(9, 7) = "兵"
start
End Sub
Private Sub as_Click()
MsgBox ("chinese-chess 中國象棋  Copyright (C) 2007-2011  Yu-Han Cheng 鄭郁翰" & vbCrLf & _
        "This program comes with ABSOLUTELY NO WARRANTY" & vbCrLf & _
        "This is free software, and you are welcome to redistribute it" & vbCrLf & _
        "GPLv3 see <http://www.gnu.org/licenses/>")

End Sub
Private Sub rego_Click()
afresh
La
End Sub
