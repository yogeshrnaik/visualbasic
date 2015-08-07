VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Game Of Digits"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Exit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   260
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "The Game of Digits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   259
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton Help 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&How To Play?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   258
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton No_of_Moves 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Squares Covered = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   257
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton NewGame 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&New Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   256
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   255
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   255
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   254
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   254
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   253
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   253
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   252
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   252
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   251
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   251
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   250
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   250
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   249
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   249
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   248
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   248
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   247
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   247
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   246
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   246
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   245
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   245
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   244
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   244
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   243
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   243
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   242
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   242
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   241
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   241
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   240
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   240
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   239
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   239
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   238
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   238
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   237
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   237
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   236
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   236
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   235
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   235
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   234
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   234
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   233
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   233
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   232
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   232
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   231
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   231
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   230
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   230
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   229
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   228
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   228
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   227
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   227
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   226
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   226
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   225
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   224
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   224
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   223
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   223
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   222
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   222
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   221
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   221
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   220
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   220
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   219
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   219
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   218
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   217
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   217
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   216
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   216
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   215
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   215
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   214
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   214
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   213
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   212
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   212
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   211
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   211
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   210
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   210
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   209
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   209
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   208
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   208
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   207
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   207
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   206
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   206
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   205
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   205
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   204
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   204
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   203
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   203
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   202
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   202
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   201
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   200
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   200
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   199
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   199
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   198
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   198
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   197
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   197
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   196
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   195
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   194
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   193
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   192
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   191
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   191
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   190
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   190
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   189
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   189
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   188
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   187
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   187
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   186
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   186
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   185
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   184
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   184
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   183
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   182
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   182
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   181
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   181
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   180
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   180
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   179
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   179
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   178
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   177
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   176
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   176
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   175
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   175
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   174
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   174
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   173
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   173
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   172
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   172
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   171
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   170
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   170
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   169
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   169
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   168
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   168
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   167
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   166
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   165
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   164
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   163
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   162
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   161
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   160
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   159
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   158
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   157
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   156
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   155
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   154
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   153
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   152
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   151
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   150
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   149
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   148
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   147
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   146
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   145
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   145
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   144
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   143
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   143
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   142
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   141
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   140
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   139
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   138
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   137
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   136
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   135
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   134
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   133
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   132
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   131
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   130
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   129
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   128
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   127
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   126
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   125
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   124
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   123
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   122
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   121
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   120
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   119
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   118
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   117
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   116
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   115
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   114
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   113
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   112
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   111
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   110
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   109
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   108
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   107
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   106
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   105
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   104
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   103
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   102
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   101
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   100
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   99
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   98
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   97
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   96
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   95
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   94
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   93
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   92
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   91
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   90
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   89
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   88
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   87
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   86
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   85
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   84
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   83
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   82
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   81
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   80
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   79
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   78
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   77
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   76
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   75
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   74
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   73
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   72
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   71
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   70
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   69
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   68
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   67
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   66
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   65
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   64
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   63
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   62
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   61
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   60
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   59
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   58
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   57
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   56
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   55
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   54
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   53
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   52
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   51
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   50
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   49
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   48
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   47
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   46
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   45
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   44
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   43
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   42
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   41
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   40
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   39
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   38
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   37
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   36
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   35
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   34
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   33
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   32
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   31
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   30
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   29
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   28
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   27
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   26
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   25
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   24
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   23
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   22
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   21
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   20
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   19
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   18
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   17
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   16
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   15
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   14
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   13
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   12
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   11
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   10
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   9
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   8
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   7
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   6
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   5
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   4
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   3
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   2
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Digits 
      Height          =   375
      Index           =   0
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim start As Integer
Dim zeros As Integer, fives As Integer
Dim num As Integer, total_moves As Integer
Private Sub Command1_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Digits_Click(Index As Integer)
    'check where is the clicked cell w.r.t. start cell
    '1)start cell is at the bottom of clicked cell
    If Digits(Index).top + 360 = Digits(start).top And Digits(Index).left = Digits(start).left Then
        'check if move is possible in top direction
        If Digits(Index).Caption <> "" Then
            If isMovePossible(Index, 0) Then
                'move start in top dir
                moveInDir Index, 0
            End If
        End If
    End If
    
    '2)start cell is on the right side of clicked cell
    If Digits(Index).left + 360 = Digits(start).left And Digits(Index).top = Digits(start).top Then
        'check if move is possible in left direction
        If Digits(Index).Caption <> "" Then
            If isMovePossible(Index, 1) Then
                'move start in left dir
                moveInDir Index, 1
            End If
        End If
    End If
    
    '3)start cell is at the top of clicked cell
    If Digits(Index).top - 360 = Digits(start).top And Digits(Index).left = Digits(start).left Then
        'check if move is possible in down direction
        If Digits(Index).Caption <> "" Then
            If isMovePossible(Index, 2) Then
                'move start in down dir
                moveInDir Index, 2
            End If
        End If
    End If
               
    '4)start cell is on the left side of clicked cell
    If Digits(Index).left - 360 = Digits(start).left And Digits(Index).top = Digits(start).top Then
        'check if move is possible in right direction
        If Digits(Index).Caption <> "" Then
            If isMovePossible(Index, 3) Then
                'move start in right dir
                moveInDir Index, 3
            End If
        End If
    End If
    
    '5)start cell is at the bottom right of clicked cell
    If Digits(Index).top + 360 = Digits(start).top And Digits(Index).left + 360 = Digits(start).left Then
        'check if move is possible in top left direction
        If Digits(Index).Caption <> "" Then
            If isMovePossible(Index, 4) Then
                'move start in top dir
                moveInDir Index, 4
            End If
        End If
    End If
    
    '6)start cell is at the top right of clicked cell
    If Digits(Index).top - 360 = Digits(start).top And Digits(Index).left + 360 = Digits(start).left Then
        'check if move is possible in top left direction
        If Digits(Index).Caption <> "" Then
            If isMovePossible(Index, 5) Then
                'move start in top dir
                moveInDir Index, 5
            End If
        End If
    End If
    
    
    '7)start cell is at the top left of clicked cell
    If Digits(Index).top - 360 = Digits(start).top And Digits(Index).left - 360 = Digits(start).left Then
        'check if move is possible in top left direction
        If Digits(Index).Caption <> "" Then
            If isMovePossible(Index, 6) Then
                'move start in top dir
                moveInDir Index, 6
            End If
        End If
    End If
    
    '8)start cell is at the bottom left of clicked cell
    If Digits(Index).top + 360 = Digits(start).top And Digits(Index).left - 360 = Digits(start).left Then
        'check if move is possible in top left direction
        If Digits(Index).Caption <> "" Then
            If isMovePossible(Index, 7) Then
                'move start in top dir
                moveInDir Index, 7
            End If
        End If
    End If
        
    If isGameOver Then
        MsgBox "GAME OVER! No More Moves Possible!! You covered " & total_moves & " squares.", , "Game Of Digits"
        NewGame_Click
    End If
End Sub

Private Sub Digits_KeyPress(Index As Integer, KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub Exit_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Form_Load()
    NewGame_Click
End Sub
Private Sub generate()
    Dim count As Integer
    zeros = 0
    fives = 0
    Randomize
    For count = 0 To 255
        num = Rnd * 100
        If num Mod 5 <> 0 Then
            Digits(count).Caption = num Mod 5
            Digits(count).FontBold = False
            Digits(count).BackColor = &HC0C0C0
        Else
            count = count - 1
        End If
        'Digits(count).Visible = True
    Next
    For count = 0 To 9
        num = (Rnd * 10000) Mod 256
        Digits(num).Caption = ""
        Digits(num).BackColor = &H404000
    Next
    
    start = (Rnd * 100000) Mod 255
    Digits(start).Caption = "S"
    Digits(start).FontBold = True
    Digits(start).BackColor = &H80C0FF   '&H80FF& (Red)
End Sub

Private Sub Help_Click()
    MsgBox "Welcome to the 'Game of Digits'. The game is played in a 16 X 16 matrix of squares. The start position is denoted by 'S'. You can move the start position horizontally or vertically or diagonally by clicking on the square that is adjacent to the starting position or by using the Num-pad Keys. Depending on the number that is there on the clicked square, the starting position is shifted that many squares horizontally or vertically or diagonally. Your objective is to cover maximum number of squares. The Game gets over if there are no moves possible.", vbInformation, "Game Of Digits"
End Sub

Private Sub Help_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub NewGame_Click()
    generate
    total_moves = 0
    No_of_Moves.Caption = "Squares Covered = 0"
End Sub

Function isMovePossible(Index As Integer, dir As Integer) As Boolean
    Dim top As Integer, left As Integer
    Dim possible As Boolean
    Dim temp As Integer, moves As Integer, count As Integer
       
    temp = start
    possible = True
    If Digits(Index).Caption <> "" And Digits(Index).Caption <> "S" Then
        moves = Val(Digits(Index).Caption)
    Else
        possible = False
        isMovePossible = False
    End If
    For count = 1 To moves
        'move in up direction
        If dir = 0 Then
            temp = temp - 16
        End If
        
        'move in left direction
        If dir = 1 Then
            temp = temp - 1
        End If
        
        'move in down direction
        If dir = 2 Then
            temp = temp + 16
        End If
        
        'move in right direction
        If dir = 3 Then
            temp = temp + 1
        End If
        
        'move in top left direction
        If dir = 4 Then
            If temp Mod 16 <> 0 Then
                temp = temp - 17
            Else
                possible = False
                isMovePossible = False
            End If
        End If
        
        'move in bottom left direction
        If dir = 5 Then
            If temp Mod 16 <> 0 Then
                temp = temp + 15
            Else
                possible = False
                isMovePossible = False
            End If
        End If
        
        'move in bottom right direction
        If dir = 6 Then
            If (temp + 1) Mod 16 <> 0 Then
                temp = temp + 17
            Else
                possible = False
                isMovePossible = False
            End If
        End If
        
        
        'move in top right direction
        If dir = 7 Then
            If (temp + 1) Mod 16 <> 0 Then
                temp = temp - 15
            Else
                possible = False
                isMovePossible = False
            End If
        End If
        
        If temp >= 0 And temp <= 255 Then
            If Digits(temp).Caption = "" Or Digits(temp).Caption = "S" Then
                possible = False
                isMovePossible = False
            End If
        Else
            possible = False
            isMovePossible = False
        End If
    Next
    
    If possible Then
        'check if move is possible in top dir (0)
        If dir = 0 Then
            top = Digits(Index).top - (360 * Val(Digits(Index).Caption))
            If (top >= 1200 And top <= 7320) Then
                isMovePossible = True
            End If
            'MsgBox top
        End If
    
        'check if move is possible in left dir (1)
        If dir = 1 Then
            left = Digits(Index).left - (360 * Val(Digits(Index).Caption))
            If (left >= 240 And left <= 6360) Then
                isMovePossible = True
            End If
            'MsgBox left
        End If
    
        'check if move is possible in down dir (2)
        If dir = 2 Then
            top = Digits(Index).top + (360 * Val(Digits(Index).Caption))
            If (top >= 1200 And top <= 7320) Then
                isMovePossible = True
            End If
            'MsgBox top
        End If
    
        'check if move is possible in right dir (3)
        If dir = 3 Then
            left = Digits(Index).left + (360 * Val(Digits(Index).Caption))
            If (left >= 240 And left <= 6360) Then
                isMovePossible = True
            End If
            'MsgBox left
        End If
        
        'check if move is possible in top left dir (4)
        If dir = 4 Then
            top = Digits(Index).top - (360 * Val(Digits(Index).Caption))
            left = Digits(Index).left - (360 * Val(Digits(Index).Caption))
            If (top >= 1200 And top <= 7320) Then
                If (left >= 240 And left <= 6000) Then
                    isMovePossible = True
                End If
            End If
            'MsgBox top
        End If
        
        'check if move is possible in bottom left dir (5)
        If dir = 5 Then
            top = Digits(Index).top + (360 * Val(Digits(Index).Caption))
            left = Digits(Index).left - (360 * Val(Digits(Index).Caption))
            If (top >= 1200 And top <= 7320) Then
                If (left >= 240 And left <= 6000) Then
                    isMovePossible = True
                End If
            End If
            'MsgBox top
        End If
        
        'check if move is possible in bottom right dir (6)
        If dir = 6 Then
            top = Digits(Index).top + (360 * Val(Digits(Index).Caption))
            left = Digits(Index).left + (360 * Val(Digits(Index).Caption))
            If (top >= 1200 And top <= 7320) Then
                If (left >= 240 And left <= 6360) Then
                    isMovePossible = True
                End If
            End If
            'MsgBox top
        End If
        
        
        'check if move is possible in top right dir (7)
        If dir = 7 Then
            top = Digits(Index).top - (360 * Val(Digits(Index).Caption))
            left = Digits(Index).left + (360 * Val(Digits(Index).Caption))
            If (top >= 1200 And top <= 7320) Then
                If (left >= 240 And left <= 6360) Then
                    isMovePossible = True
                End If
            End If
            'MsgBox top
        End If
        
    End If
End Function

Private Sub moveInDir(Index As Integer, dir As Integer)
    Dim count As Integer, moves As Integer
    moves = Val(Digits(Index).Caption)
    total_moves = total_moves + moves
    No_of_Moves.Caption = "Squares Covered = " & total_moves
    For count = 1 To moves
        Digits(start).Caption = ""
        Digits(start).BackColor = &H80FF&
        Digits(start).FontBold = False
        
        'move in up direction
        If dir = 0 Then
            start = start - 16
        End If
        
        'move in left direction
        If dir = 1 Then
            start = start - 1
        End If
        
        'move in down direction
        If dir = 2 Then
            start = start + 16
        End If
        
        'move in right direction
        If dir = 3 Then
            start = start + 1
        End If
        
        'move in top left direction
        If dir = 4 Then
            start = start - 17
        End If
        
        'move in bottom left direction
        If dir = 5 Then
            start = start + 15
        End If
        
        'move in bottom right direction
        If dir = 6 Then
            start = start + 17
        End If
        
        'move in top right direction
        If dir = 7 Then
            start = start - 15
        End If
        
        Digits(start).Caption = ""
        Digits(start).BackColor = &H80FF&
        Digits(start).FontBold = False
        
    Next
    Digits(start).Caption = "S"
    Digits(start).BackColor = &H80C0FF
    Digits(start).FontBold = True
End Sub

Function isGameOver() As Boolean
    Dim count As Integer
    
    'check if move is possible in up direction
    If (start - 16) >= 0 Then
        If Not isMovePossible(start - 16, 0) Then
           count = count + 1
        End If
    Else
        count = count + 1
    End If
    
    'check id move is possible in left direction
    'also check if the start point is at extreme left
    'start mod 16 = 0
    If (start - 1) >= 0 And start Mod 16 <> 0 Then
        If Not isMovePossible(start - 1, 1) Then
           count = count + 1
        End If
    Else
        count = count + 1
    End If
    
    'check if move is possible in down direction
    If (start + 16) <= 255 Then
        If Not isMovePossible(start + 16, 2) Then
           count = count + 1
        End If
    Else
        count = count + 1
    End If
    
    'check if move is possible in right direction
    'also check if the start point is at extreme right
    If (start + 1) <= 255 And (start + 1) Mod 16 <> 0 Then
        If Not isMovePossible(start + 1, 3) Then
           count = count + 1
        End If
    Else
        count = count + 1
    End If
    
    
    'check if move is possible in top left direction
    'also check if the start point is at extreme left
    If (start - 17) >= 0 And start Mod 16 <> 0 Then
        If Not isMovePossible(start - 17, 4) Then
           count = count + 1
        End If
    Else
        count = count + 1
    End If
    
    'check if move is possible in bottom left direction
    'also check if the start point is at extreme left
    If (start + 15) <= 255 And start Mod 16 <> 0 Then
        If Not isMovePossible(start + 15, 5) Then
           count = count + 1
        End If
    Else
        count = count + 1
    End If
    
    
    'check if move is possible in bottom right direction
    'also check if the start point is at extreme right
    If (start + 17) <= 255 And (start + 1) Mod 16 <> 0 Then
        If Not isMovePossible(start + 17, 6) Then
           count = count + 1
        End If
    Else
        count = count + 1
    End If
    
    'check if move is possible in top right direction
    'also check if the start point is at extreme right
    If (start - 15) >= 0 And (start + 1) Mod 16 <> 0 Then
        If Not isMovePossible(start - 15, 7) Then
           count = count + 1
        End If
    Else
        count = count + 1
    End If
    
    
    If count >= 8 Then
        isGameOver = True
    Else
        isGameOver = False
    End If
    
    'MsgBox count
End Function

Function Handle_KeyPressed(KeyAscii)
    '1 pressed - move in bottom left direction
    If KeyAscii = 49 And start + 15 <= 255 Then
        Digits_Click (start + 15)
    End If
    '2 pressed - move in down direction
    If KeyAscii = 50 And start + 16 <= 255 Then
        Digits_Click (start + 16)
    End If
    '3 pressed - move in bottom right direction
    If KeyAscii = 51 And start + 17 <= 255 Then
        Digits_Click (start + 17)
    End If
    '4 pressed - move in left direction
    If KeyAscii = 52 And start - 1 >= 0 Then
        Digits_Click (start - 1)
    End If
    '6 pressed - move in right direction
    If KeyAscii = 54 And start + 1 <= 255 Then
        Digits_Click (start + 1)
    End If
    '7 pressed - move in top left direction
    If KeyAscii = 55 And start - 17 >= 0 Then
        Digits_Click (start - 17)
    End If
    '8 pressed - move in up direction
    If KeyAscii = 56 And start - 16 >= 0 Then
        Digits_Click (start - 16)
    End If
    '9 pressed - move in top right direction
    If KeyAscii = 57 And start - 15 >= 0 Then
        Digits_Click (start - 15)
    End If
End Function

Private Sub NewGame_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub
Private Sub No_of_Moves_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub
