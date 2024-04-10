VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "腾讯会议辅助"
   ClientHeight    =   11805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11805
   ScaleWidth      =   17985
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command14 
      Caption         =   "清空日志"
      Height          =   495
      Left            =   9600
      TabIndex        =   110
      Top             =   11160
      Width           =   8175
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   16440
      Top             =   8640
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15960
      Top             =   8640
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15480
      Top             =   8640
   End
   Begin VB.Timer Timer17 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15000
      Top             =   8640
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14520
      Top             =   8640
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   16440
      Top             =   8160
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15960
      Top             =   8160
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15480
      Top             =   8160
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15000
      Top             =   8160
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14520
      Top             =   8160
   End
   Begin VB.CheckBox Check1 
      Caption         =   "下课自动重启进程(只支持默认安装路径)"
      Height          =   300
      Left            =   4080
      TabIndex        =   109
      Top             =   5880
      Width           =   4455
   End
   Begin VB.TextBox Text41 
      Height          =   270
      Left            =   7440
      TabIndex        =   108
      Text            =   "0"
      Top             =   12600
      Width           =   1215
   End
   Begin VB.TextBox Text40 
      Height          =   270
      Left            =   7440
      TabIndex        =   107
      Text            =   "0"
      Top             =   13080
      Width           =   1215
   End
   Begin VB.TextBox Text39 
      Height          =   270
      Left            =   7440
      TabIndex        =   106
      Text            =   "0"
      Top             =   13560
      Width           =   1215
   End
   Begin VB.TextBox Text38 
      Height          =   270
      Left            =   7440
      TabIndex        =   105
      Text            =   "0"
      Top             =   14040
      Width           =   1215
   End
   Begin VB.TextBox Text37 
      Height          =   270
      Left            =   7440
      TabIndex        =   104
      Text            =   "0"
      Top             =   14520
      Width           =   1215
   End
   Begin VB.TextBox Text36 
      Height          =   270
      Left            =   7440
      TabIndex        =   103
      Text            =   "0"
      Top             =   15000
      Width           =   1215
   End
   Begin VB.TextBox Text35 
      Height          =   270
      Left            =   7440
      TabIndex        =   102
      Text            =   "0"
      Top             =   15480
      Width           =   1215
   End
   Begin VB.TextBox Text34 
      Height          =   270
      Left            =   7440
      TabIndex        =   101
      Text            =   "0"
      Top             =   15960
      Width           =   1215
   End
   Begin VB.TextBox Text33 
      Height          =   270
      Left            =   11640
      TabIndex        =   100
      Text            =   "0"
      Top             =   15840
      Width           =   1215
   End
   Begin VB.TextBox Text32 
      Height          =   270
      Left            =   11640
      TabIndex        =   99
      Text            =   "0"
      Top             =   16320
      Width           =   1215
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   16440
      Top             =   7080
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15960
      Top             =   7080
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15480
      Top             =   7080
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15000
      Top             =   7080
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14520
      Top             =   7080
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   16440
      Top             =   6600
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15960
      Top             =   6600
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15480
      Top             =   6600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14520
      Top             =   6600
   End
   Begin VB.ComboBox Combo10 
      Height          =   300
      Left            =   7200
      TabIndex        =   98
      Text            =   "空"
      Top             =   11280
      Width           =   1575
   End
   Begin VB.ComboBox Combo9 
      Height          =   300
      Left            =   7200
      TabIndex        =   97
      Text            =   "空"
      Top             =   10800
      Width           =   1575
   End
   Begin VB.ComboBox Combo8 
      Height          =   300
      Left            =   7200
      TabIndex        =   96
      Text            =   "空"
      Top             =   10320
      Width           =   1575
   End
   Begin VB.ComboBox Combo7 
      Height          =   300
      Left            =   7200
      TabIndex        =   95
      Text            =   "空"
      Top             =   9840
      Width           =   1575
   End
   Begin VB.ComboBox Combo6 
      Height          =   300
      Left            =   7200
      TabIndex        =   94
      Text            =   "空"
      Top             =   9360
      Width           =   1575
   End
   Begin VB.ComboBox Combo5 
      Height          =   300
      Left            =   7200
      TabIndex        =   93
      Text            =   "空"
      Top             =   8880
      Width           =   1575
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   7200
      TabIndex        =   92
      Text            =   "空"
      Top             =   8400
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   7200
      TabIndex        =   91
      Text            =   "空"
      Top             =   7920
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   7200
      TabIndex        =   90
      Text            =   "空"
      Top             =   7440
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   7200
      TabIndex        =   89
      Text            =   "空"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox Text31 
      Height          =   270
      Left            =   4320
      TabIndex        =   77
      Text            =   "08:45:00"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text30 
      Height          =   270
      Left            =   4320
      TabIndex        =   76
      Text            =   "09:45:00"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text29 
      Height          =   270
      Left            =   4320
      TabIndex        =   75
      Text            =   "10:45:00"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox Text28 
      Height          =   270
      Left            =   4320
      TabIndex        =   74
      Text            =   "11:45:00"
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox Text27 
      Height          =   270
      Left            =   4320
      TabIndex        =   73
      Text            =   "15:15:00"
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox Text26 
      Height          =   270
      Left            =   4320
      TabIndex        =   72
      Text            =   "16:15:00"
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text25 
      Height          =   270
      Left            =   4320
      TabIndex        =   71
      Text            =   "17:15:00"
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox Text24 
      Height          =   270
      Left            =   4320
      TabIndex        =   70
      Text            =   "20:10:00"
      Top             =   10320
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Height          =   270
      Left            =   4320
      TabIndex        =   69
      Text            =   "21:10:00"
      Top             =   10800
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      Height          =   270
      Left            =   4320
      TabIndex        =   68
      Text            =   "22:10:00"
      Top             =   11280
      Width           =   1215
   End
   Begin VB.TextBox Text21 
      Height          =   270
      Left            =   1440
      TabIndex        =   55
      Text            =   "21:20:00"
      Top             =   11280
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Height          =   270
      Left            =   1440
      TabIndex        =   54
      Text            =   "20:20:00"
      Top             =   10800
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      Height          =   270
      Left            =   1440
      TabIndex        =   53
      Text            =   "19:20:00"
      Top             =   10320
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   270
      Left            =   1440
      TabIndex        =   52
      Text            =   "16:30:00"
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   270
      Left            =   1440
      TabIndex        =   51
      Text            =   "15:30:00"
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   270
      Left            =   1440
      TabIndex        =   50
      Text            =   "14:30:00"
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Left            =   1440
      TabIndex        =   49
      Text            =   "11:00:00"
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Left            =   1440
      TabIndex        =   48
      Text            =   "10:00:00"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   1440
      TabIndex        =   47
      Text            =   "09:00:00"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15000
      Top             =   6600
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   1440
      TabIndex        =   34
      Text            =   "08:00:00"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Timer Timer 
      Left            =   2880
      Top             =   5880
   End
   Begin VB.CommandButton Command13 
      Caption         =   "强制关闭所有腾讯会议程序"
      Height          =   615
      Left            =   6840
      TabIndex        =   33
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton Command12 
      Caption         =   "加载会议号设置"
      Height          =   615
      Left            =   3600
      TabIndex        =   32
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      Caption         =   "保存会议号设置"
      Height          =   615
      Left            =   240
      TabIndex        =   31
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "手动加入政治"
      Height          =   375
      Left            =   7320
      TabIndex        =   30
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "手动加入化学"
      Height          =   375
      Left            =   7320
      TabIndex        =   29
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "手动加入地理"
      Height          =   375
      Left            =   7320
      TabIndex        =   28
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "手动加入生物"
      Height          =   375
      Left            =   7320
      TabIndex        =   27
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "手动加入物理"
      Height          =   375
      Left            =   7320
      TabIndex        =   26
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "手动加入历史"
      Height          =   375
      Left            =   7320
      TabIndex        =   25
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "手动加入外语"
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "手动加入数学"
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   2160
      TabIndex        =   22
      Top             =   4680
      Width           =   5055
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   2160
      TabIndex        =   21
      Top             =   4200
      Width           =   5055
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   2160
      TabIndex        =   20
      Top             =   3720
      Width           =   5055
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   2160
      TabIndex        =   19
      Top             =   3240
      Width           =   5055
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   2160
      TabIndex        =   18
      Top             =   2760
      Width           =   5055
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   2160
      TabIndex        =   17
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   2160
      TabIndex        =   16
      Top             =   1800
      Width           =   5055
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "使用默认安装路径"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   10695
      Left            =   9600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "手动加入语文"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label45 
      Caption         =   "第一节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   88
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label44 
      Caption         =   "第二节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   87
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label43 
      Caption         =   "第三节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   86
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label42 
      Caption         =   "第四节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   85
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label41 
      Caption         =   "第五节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   84
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label Label40 
      Caption         =   "第六节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   83
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label Label39 
      Caption         =   "第七节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   82
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label Label38 
      Caption         =   "第八节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   81
      Top             =   10320
      Width           =   975
   End
   Begin VB.Label Label37 
      Caption         =   "第九节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   80
      Top             =   10800
      Width           =   975
   End
   Begin VB.Label Label36 
      Caption         =   "第十节课:"
      Height          =   255
      Left            =   6000
      TabIndex        =   79
      Top             =   11280
      Width           =   975
   End
   Begin VB.Label Label35 
      Caption         =   "课表:"
      Height          =   255
      Left            =   6000
      TabIndex        =   78
      Top             =   6480
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   5760
      Y1              =   6240
      Y2              =   11520
   End
   Begin VB.Label Label34 
      Caption         =   "第一节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   67
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label33 
      Caption         =   "第二节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   66
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label32 
      Caption         =   "第三节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   65
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label31 
      Caption         =   "第四节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   64
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label30 
      Caption         =   "第五节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   63
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "第六节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   62
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "第七节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   61
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label Label27 
      Caption         =   "第八节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   60
      Top             =   10320
      Width           =   975
   End
   Begin VB.Label Label26 
      Caption         =   "第九节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   59
      Top             =   10800
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "第十节课:"
      Height          =   255
      Left            =   3120
      TabIndex        =   58
      Top             =   11280
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "下课时间 :"
      Height          =   255
      Left            =   3120
      TabIndex        =   57
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   6240
      Y2              =   11520
   End
   Begin VB.Label Label23 
      Caption         =   "上课时间:"
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label22 
      Caption         =   "第十节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   11280
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "第九节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   45
      Top             =   10800
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "第八节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   10320
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "第七节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "第六节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "第五节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "第四节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "第三节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "第二节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "第一节课:"
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   255
      Left            =   1440
      TabIndex        =   36
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "当前时间:"
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "政治会议号"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "化学会议号"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "地理会议号"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "生物会议号"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "物理会议号"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "历史会议号"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "外语会议号"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "数学会议号 "
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "语文会议号"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "腾讯会议主程序路径"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
              Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
              Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
              Private Type SECURITY_ATTRIBUTES
                              nLength   As Long
                              lpSecurityDescriptor   As Long
                              bInheritHandle   As Long
              End Type
              Private Type STARTUPINFO
                              cb   As Long
                              lpReserved   As String
                              lpDesktop   As String
                              lpTitle   As String
                              dwX   As Long
                              dwY   As Long
                              dwXSize   As Long
                              dwYSize   As Long
                              dwXCountChars   As Long
                              dwYCountChars   As Long
                              dwFillAttribute   As Long
                              dwFlags   As Long
                              wShowWindow   As Integer
                              cbReserved2   As Integer
                              lpReserved2   As Long
                              hStdInput   As Long
                              hStdOutput   As Long
                              hStdError   As Long
              End Type
              Private Type PROCESS_INFORMATION
                              hProcess   As Long
                              hThread   As Long
                              dwProcessId   As Long
                              dwThreadId   As Long
              End Type
              Private Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) As Long
              Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
              Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
              Private Const NORMAL_PRIORITY_CLASS = &H20
              Private Const STARTF_USESTDHANDLES = &H100
              Private Const STARTF_USESHOWWINDOW = &H1
              Private Const SW_HIDE = 0
              Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
              Private Function ExecuteCommandLineOutput(CommandLine As String, Optional BufferSize As Long = 256, Optional TimeOut As Long) As String
                              Dim Proc     As PROCESS_INFORMATION
                              Dim Start     As STARTUPINFO
                              Dim SA     As SECURITY_ATTRIBUTES
                              Dim hReadPipe     As Long
                              Dim hWritePipe     As Long
                              Dim lBytesRead     As Long
                              Dim sBuffer     As String
                              If VBA.Len(CommandLine) > 0 Then
                                    SA.nLength = Len(SA)
                                    'SA.nLength   =   vba.Len(sa)
                                    SA.bInheritHandle = 1&
                                    SA.lpSecurityDescriptor = 0&
                                    If CreatePipe(hReadPipe, hWritePipe, SA, 0) > 0 Then
                                          Start.cb = Len(Start)
                                          Start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
                                          Start.hStdOutput = hWritePipe
                                          Start.hStdError = hWritePipe
                                          Start.wShowWindow = SW_HIDE
                                          If CreateProcessA(0&, CommandLine, SA, SA, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, Proc) = 1 Then
                                                CloseHandle hWritePipe
                                                sBuffer = VBA.String(BufferSize, VBA.Chr(0))
                                                If TimeOut > 0 Then
                                                      Dim BeginTime     As Date
                                                      BeginTime = VBA.Now
                                                End If
                                                Do Until ReadFile(hReadPipe, sBuffer, BufferSize, lBytesRead, 0&) = 0
                                                      DoEvents
                                                      If TimeOut > 0 Then
                                                            If VBA.DateDiff("s", BeginTime, VBA.Now) > TimeOut Then
                                                                  ExecuteCommandLineOutput = "Timeout"
                                                                  Exit Do
                                                            End If
                                                      End If
                                                      ExecuteCommandLineOutput = ExecuteCommandLineOutput & VBA.Left(sBuffer, lBytesRead)
                                                Loop
                                               CloseHandle Proc.hProcess
                                                CloseHandle Proc.hThread
                                                CloseHandle hReadPipe
                                          Else
                                              ExecuteCommandLineOutput = "File   or   command   not   found"
                                        End If
                                    Else
                                        ExecuteCommandLineOutput = "CreatePipe   failed.   Error:   " & Err.LastDllError & "."
                                    End If
                            End If
              End Function

Private Sub Check1_Click()

If Form1.Check1.Value = Checked Then Timer11.Enabled = True
If Form1.Check1.Value = Checked Then Timer12.Enabled = True
If Form1.Check1.Value = Checked Then Timer13.Enabled = True
If Form1.Check1.Value = Checked Then Timer14.Enabled = True
If Form1.Check1.Value = Checked Then Timer15.Enabled = True
If Form1.Check1.Value = Checked Then Timer16.Enabled = True
If Form1.Check1.Value = Checked Then Timer17.Enabled = True
If Form1.Check1.Value = Checked Then Timer18.Enabled = True
If Form1.Check1.Value = Checked Then Timer19.Enabled = True
If Form1.Check1.Value = Checked Then Timer20.Enabled = True


If Form1.Check1.Value = Unchecked Then Timer11.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer12.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer13.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer14.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer15.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer16.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer17.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer18.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer19.Enabled = False
If Form1.Check1.Value = Unchecked Then Timer20.Enabled = False
End Sub

Private Sub Combo1_Click()
Select Case Combo1.Text
Case "空"
Timer1.Enabled = False
Case "语文"
Text41.Text = Text2.Text
Timer1.Enabled = True
Case "数学"
Text41.Text = Text4.Text
Timer1.Enabled = True
Case "外语"
Text41.Text = Text5.Text
Timer1.Enabled = True
Case "历史"
Text41.Text = Text6.Text
Timer1.Enabled = True
Case "物理"
Text41.Text = Text7.Text
Timer1.Enabled = True
Case "生物"
Text41.Text = Text8.Text
Timer1.Enabled = True
Case "地理"
Text41.Text = Text9.Text
Timer1.Enabled = True
Case "化学"
Text41.Text = Text10.Text
Timer1.Enabled = True
Case "政治"
Text41.Text = Text11.Text
Timer1.Enabled = True

End Select
End Sub



Private Sub Combo10_Click()
Select Case Combo10.Text
Case "空"
Timer10.Enabled = False
Case "语文"
Text32.Text = Text2.Text
Timer10.Enabled = True
Case "数学"
Text32.Text = Text4.Text
Timer10.Enabled = True
Case "外语"
Text32.Text = Text5.Text
Timer10.Enabled = True
Case "历史"
Text32.Text = Text6.Text
Timer10.Enabled = True
Case "物理"
Text32.Text = Text7.Text
Timer10.Enabled = True
Case "生物"
Text32.Text = Text8.Text
Timer10.Enabled = True
Case "地理"
Text32.Text = Text9.Text
Timer10.Enabled = True
Case "化学"
Text32.Text = Text10.Text
Timer10.Enabled = True
Case "政治"
Text32.Text = Text11.Text
Timer10.Enabled = True

End Select
End Sub

Private Sub Combo2_Click()
Select Case Combo2.Text
Case "空"
Timer2.Enabled = False
Case "语文"
Text40.Text = Text2.Text
Timer2.Enabled = True
Case "数学"
Text40.Text = Text4.Text
Timer2.Enabled = True
Case "外语"
Text40.Text = Text5.Text
Timer2.Enabled = True
Case "历史"
Text40.Text = Text6.Text
Timer2.Enabled = True
Case "物理"
Text40.Text = Text7.Text
Timer2.Enabled = True
Case "生物"
Text40.Text = Text8.Text
Timer2.Enabled = True
Case "地理"
Text40.Text = Text9.Text
Timer2.Enabled = True
Case "化学"
Text40.Text = Text10.Text
Timer2.Enabled = True
Case "政治"
Text40.Text = Text11.Text
Timer2.Enabled = True

End Select
End Sub



Private Sub Combo3_Click()
Select Case Combo3.Text
Case "空"
Timer3.Enabled = False
Case "语文"
Text39.Text = Text2.Text
Timer3.Enabled = True
Case "数学"
Text39.Text = Text4.Text
Timer3.Enabled = True
Case "外语"
Text39.Text = Text5.Text
Timer3.Enabled = True
Case "历史"
Text39.Text = Text6.Text
Timer3.Enabled = True
Case "物理"
Text39.Text = Text7.Text
Timer3.Enabled = True
Case "生物"
Text39.Text = Text8.Text
Timer3.Enabled = True
Case "地理"
Text39.Text = Text9.Text
Timer3.Enabled = True
Case "化学"
Text39.Text = Text10.Text
Timer3.Enabled = True
Case "政治"
Text39.Text = Text11.Text
Timer3.Enabled = True

End Select
End Sub

Private Sub Combo4_Click()
Select Case Combo4.Text
Case "空"
Timer4.Enabled = False
Case "语文"
Text38.Text = Text2.Text
Timer4.Enabled = True
Case "数学"
Text38.Text = Text4.Text
Timer4.Enabled = True
Case "外语"
Text38.Text = Text5.Text
Timer4.Enabled = True
Case "历史"
Text38.Text = Text6.Text
Timer4.Enabled = True
Case "物理"
Text38.Text = Text7.Text
Timer4.Enabled = True
Case "生物"
Text38.Text = Text8.Text
Timer4.Enabled = True
Case "地理"
Text38.Text = Text9.Text
Timer4.Enabled = True
Case "化学"
Text38.Text = Text10.Text
Timer4.Enabled = True
Case "政治"
Text38.Text = Text11.Text
Timer4.Enabled = True

End Select
End Sub

Private Sub Combo5_Click()
Select Case Combo5.Text
Case "空"
Timer5.Enabled = False
Case "语文"
Text37.Text = Text2.Text
Timer5.Enabled = True
Case "数学"
Text37.Text = Text4.Text
Timer5.Enabled = True
Case "外语"
Text37.Text = Text5.Text
Timer5.Enabled = True
Case "历史"
Text37.Text = Text6.Text
Timer5.Enabled = True
Case "物理"
Text37.Text = Text7.Text
Timer5.Enabled = True
Case "生物"
Text37.Text = Text8.Text
Timer5.Enabled = True
Case "地理"
Text37.Text = Text9.Text
Timer5.Enabled = True
Case "化学"
Text37.Text = Text10.Text
Timer5.Enabled = True
Case "政治"
Text37.Text = Text11.Text
Timer5.Enabled = True

End Select
End Sub

Private Sub Combo6_Click()
Select Case Combo6.Text
Case "空"
Timer6.Enabled = False
Case "语文"
Text36.Text = Text2.Text
Timer6.Enabled = True
Case "数学"
Text36.Text = Text4.Text
Timer6.Enabled = True
Case "外语"
Text36.Text = Text5.Text
Timer6.Enabled = True
Case "历史"
Text36.Text = Text6.Text
Timer6.Enabled = True
Case "物理"
Text36.Text = Text7.Text
Timer6.Enabled = True
Case "生物"
Text36.Text = Text8.Text
Timer6.Enabled = True
Case "地理"
Text36.Text = Text9.Text
Timer6.Enabled = True
Case "化学"
Text36.Text = Text10.Text
Timer6.Enabled = True
Case "政治"
Text36.Text = Text11.Text
Timer6.Enabled = True

End Select
End Sub

Private Sub Combo7_Click()
Select Case Combo7.Text
Case "空"
Timer7.Enabled = False
Case "语文"
Text35.Text = Text2.Text
Timer7.Enabled = True
Case "数学"
Text35.Text = Text4.Text
Timer7.Enabled = True
Case "外语"
Text35.Text = Text5.Text
Timer7.Enabled = True
Case "历史"
Text35.Text = Text6.Text
Timer7.Enabled = True
Case "物理"
Text35.Text = Text7.Text
Timer7.Enabled = True
Case "生物"
Text35.Text = Text8.Text
Timer7.Enabled = True
Case "地理"
Text35.Text = Text9.Text
Timer7.Enabled = True
Case "化学"
Text35.Text = Text10.Text
Timer7.Enabled = True
Case "政治"
Text35.Text = Text11.Text
Timer7.Enabled = True

End Select
End Sub

Private Sub Combo8_Click()
Select Case Combo8.Text
Case "空"
Timer8.Enabled = False
Case "语文"
Text34.Text = Text2.Text
Timer8.Enabled = True
Case "数学"
Text34.Text = Text4.Text
Timer8.Enabled = True
Case "外语"
Text34.Text = Text5.Text
Timer8.Enabled = True
Case "历史"
Text34.Text = Text6.Text
Timer8.Enabled = True
Case "物理"
Text34.Text = Text7.Text
Timer8.Enabled = True
Case "生物"
Text34.Text = Text8.Text
Timer8.Enabled = True
Case "地理"
Text34.Text = Text9.Text
Timer8.Enabled = True
Case "化学"
Text34.Text = Text10.Text
Timer8.Enabled = True
Case "政治"
Text34.Text = Text11.Text
Timer8.Enabled = True

End Select
End Sub

Private Sub Combo9_Click()
Select Case Combo9.Text
Case "空"
Timer9.Enabled = False
Case "语文"
Text33.Text = Text2.Text
Timer9.Enabled = True
Case "数学"
Text33.Text = Text4.Text
Timer9.Enabled = True
Case "外语"
Text33.Text = Text5.Text
Timer9.Enabled = True
Case "历史"
Text33.Text = Text6.Text
Timer9.Enabled = True
Case "物理"
Text33.Text = Text7.Text
Timer9.Enabled = True
Case "生物"
Text33.Text = Text8.Text
Timer9.Enabled = True
Case "地理"
Text33.Text = Text9.Text
Timer9.Enabled = True
Case "化学"
Text33.Text = Text10.Text
Timer9.Enabled = True
Case "政治"
Text33.Text = Text11.Text
Timer9.Enabled = True

End Select
End Sub

Private Sub Command1_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入语文会议号中-----" & vbCrLf
DoEvents
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text2.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Command10_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入政治会议号中-----" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text11.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Command11_Click()
SaveSetting "TxhyAuto", "exe", "exe", Text3.Text
SaveSetting "TxhyAuto", "number", "yw", Text2.Text
SaveSetting "TxhyAuto", "number", "sx", Text4.Text
SaveSetting "TxhyAuto", "number", "wy", Text5.Text
SaveSetting "TxhyAuto", "number", "ls", Text6.Text
SaveSetting "TxhyAuto", "number", "wl", Text7.Text
SaveSetting "TxhyAuto", "number", "sw", Text8.Text
SaveSetting "TxhyAuto", "number", "dl", Text9.Text
SaveSetting "TxhyAuto", "number", "hx", Text10.Text
SaveSetting "TxhyAuto", "number", "zz", Text11.Text

SaveSetting "TxhyAuto", "classtime", "1", Text12.Text
SaveSetting "TxhyAuto", "classtime", "2", Text13.Text
SaveSetting "TxhyAuto", "classtime", "3", Text14.Text
SaveSetting "TxhyAuto", "classtime", "4", Text15.Text
SaveSetting "TxhyAuto", "classtime", "5", Text16.Text
SaveSetting "TxhyAuto", "classtime", "6", Text17.Text
SaveSetting "TxhyAuto", "classtime", "7", Text18.Text
SaveSetting "TxhyAuto", "classtime", "8", Text19.Text
SaveSetting "TxhyAuto", "classtime", "9", Text20.Text
SaveSetting "TxhyAuto", "classtime", "10", Text21.Text

End Sub

Private Sub Command12_Click()
Text3.Text = GetSetting("TxhyAuto", "exe", "exe")
Text2.Text = GetSetting("TxhyAuto", "number", "yw")
Text4.Text = GetSetting("TxhyAuto", "number", "sx")
Text5.Text = GetSetting("TxhyAuto", "number", "wy")
Text6.Text = GetSetting("TxhyAuto", "number", "ls")
Text7.Text = GetSetting("TxhyAuto", "number", "wl")
Text8.Text = GetSetting("TxhyAuto", "number", "sw")
Text9.Text = GetSetting("TxhyAuto", "number", "dl")
Text10.Text = GetSetting("TxhyAuto", "number", "hx")
Text11.Text = GetSetting("TxhyAuto", "number", "zz")

Text12.Text = GetSetting("TxhyAuto", "classtime", "1")
Text13.Text = GetSetting("TxhyAuto", "classtime", "2")
Text14.Text = GetSetting("TxhyAuto", "classtime", "3")
Text15.Text = GetSetting("TxhyAuto", "classtime", "4")
Text16.Text = GetSetting("TxhyAuto", "classtime", "5")
Text17.Text = GetSetting("TxhyAuto", "classtime", "6")
Text18.Text = GetSetting("TxhyAuto", "classtime", "7")
Text19.Text = GetSetting("TxhyAuto", "classtime", "8")
Text20.Text = GetSetting("TxhyAuto", "classtime", "9")
Text21.Text = GetSetting("TxhyAuto", "classtime", "10")


End Sub

Private Sub Command13_Click()
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
End Sub

Private Sub Command14_Click()
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Text3.Text = "C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe"
End Sub


Private Sub Command3_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入数学会议号中-----" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text4.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Command4_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入外语会议号中-----" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text5.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Command5_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入历史会议号中-----" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text6.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Command6_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入物理会议号中-----" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text7.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Command7_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入生物会议号中-----" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text8.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Command8_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入地理会议号中-----" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text9.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Command9_Click()
DoEvents
Text1.Text = Text1.Text + "-----加入化学会议号中-----" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text10.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
End Sub

Private Sub Form_Load()
Text3.Text = GetSetting("TxhyAuto", "exe", "exe")
Text2.Text = GetSetting("TxhyAuto", "number", "yw")
Text4.Text = GetSetting("TxhyAuto", "number", "sx")
Text5.Text = GetSetting("TxhyAuto", "number", "wy")
Text6.Text = GetSetting("TxhyAuto", "number", "ls")
Text7.Text = GetSetting("TxhyAuto", "number", "wl")
Text8.Text = GetSetting("TxhyAuto", "number", "sw")
Text9.Text = GetSetting("TxhyAuto", "number", "dl")
Text10.Text = GetSetting("TxhyAuto", "number", "hx")
Text11.Text = GetSetting("TxhyAuto", "number", "zz")


Text12.Text = GetSetting("TxhyAuto", "classtime", "1")
Text13.Text = GetSetting("TxhyAuto", "classtime", "2")
Text14.Text = GetSetting("TxhyAuto", "classtime", "3")
Text15.Text = GetSetting("TxhyAuto", "classtime", "4")
Text16.Text = GetSetting("TxhyAuto", "classtime", "5")
Text17.Text = GetSetting("TxhyAuto", "classtime", "6")
Text18.Text = GetSetting("TxhyAuto", "classtime", "7")
Text19.Text = GetSetting("TxhyAuto", "classtime", "8")
Text20.Text = GetSetting("TxhyAuto", "classtime", "9")
Text21.Text = GetSetting("TxhyAuto", "classtime", "10")

Text1.Text = "此程序仅供学习参考使用,不得用于商业用途!" & vbCrLf
Text1.Text = Text1.Text + "版本: V1.1" & vbCrLf
Text1.Text = Text1.Text + "纪念高中时代特别版" & vbCrLf
Text1.Text = Text1.Text + "Devloped By Phtcloud_Dev" & vbCrLf & vbCrLf
Text1.Text = Text1.Text + "建议使用手机号登录 (避免自动结束重启软件跳登录页面)" & vbCrLf & vbCrLf
Text1.Text = Text1.Text + "使用前请在腾讯会议设置里开启 入会时使用电脑音频" & vbCrLf & vbCrLf
Text1.Text = Text1.Text + "此软件使用Microsoft Visual Basic编译,理论无需额外dll组件" & vbCrLf & vbCrLf

Combo1.AddItem "空"
Combo1.AddItem "语文"
Combo1.AddItem "数学"
Combo1.AddItem "外语"
Combo1.AddItem "历史"
Combo1.AddItem "物理"
Combo1.AddItem "生物"
Combo1.AddItem "地理"
Combo1.AddItem "化学"
Combo1.AddItem "政治"

Combo2.AddItem "空"
Combo2.AddItem "语文"
Combo2.AddItem "数学"
Combo2.AddItem "外语"
Combo2.AddItem "历史"
Combo2.AddItem "物理"
Combo2.AddItem "生物"
Combo2.AddItem "地理"
Combo2.AddItem "化学"
Combo2.AddItem "政治"

Combo3.AddItem "空"
Combo3.AddItem "语文"
Combo3.AddItem "数学"
Combo3.AddItem "外语"
Combo3.AddItem "历史"
Combo3.AddItem "物理"
Combo3.AddItem "生物"
Combo3.AddItem "地理"
Combo3.AddItem "化学"
Combo3.AddItem "政治"

Combo4.AddItem "空"
Combo4.AddItem "语文"
Combo4.AddItem "数学"
Combo4.AddItem "外语"
Combo4.AddItem "历史"
Combo4.AddItem "物理"
Combo4.AddItem "生物"
Combo4.AddItem "地理"
Combo4.AddItem "化学"
Combo4.AddItem "政治"

Combo5.AddItem "空"
Combo5.AddItem "语文"
Combo5.AddItem "数学"
Combo5.AddItem "外语"
Combo5.AddItem "历史"
Combo5.AddItem "物理"
Combo5.AddItem "生物"
Combo5.AddItem "地理"
Combo5.AddItem "化学"
Combo5.AddItem "政治"

Combo6.AddItem "空"
Combo6.AddItem "语文"
Combo6.AddItem "数学"
Combo6.AddItem "外语"
Combo6.AddItem "历史"
Combo6.AddItem "物理"
Combo6.AddItem "生物"
Combo6.AddItem "地理"
Combo6.AddItem "化学"
Combo6.AddItem "政治"

Combo7.AddItem "空"
Combo7.AddItem "语文"
Combo7.AddItem "数学"
Combo7.AddItem "外语"
Combo7.AddItem "历史"
Combo7.AddItem "物理"
Combo7.AddItem "生物"
Combo7.AddItem "地理"
Combo7.AddItem "化学"
Combo7.AddItem "政治"

Combo8.AddItem "空"
Combo8.AddItem "语文"
Combo8.AddItem "数学"
Combo8.AddItem "外语"
Combo8.AddItem "历史"
Combo8.AddItem "物理"
Combo8.AddItem "生物"
Combo8.AddItem "地理"
Combo8.AddItem "化学"
Combo8.AddItem "政治"

Combo9.AddItem "空"
Combo9.AddItem "语文"
Combo9.AddItem "数学"
Combo9.AddItem "外语"
Combo9.AddItem "历史"
Combo9.AddItem "物理"
Combo9.AddItem "生物"
Combo9.AddItem "地理"
Combo9.AddItem "化学"
Combo9.AddItem "政治"

Combo10.AddItem "空"
Combo10.AddItem "语文"
Combo10.AddItem "数学"
Combo10.AddItem "外语"
Combo10.AddItem "历史"
Combo10.AddItem "物理"
Combo10.AddItem "生物"
Combo10.AddItem "地理"
Combo10.AddItem "化学"
Combo10.AddItem "政治"


Timer.Enabled = True
Timer.Interval = 1
Timer1.Enabled = True
Timer1.Interval = 1
Timer2.Enabled = True
Timer2.Interval = 1



End Sub



Private Sub Timer_Timer()
DoEvents
Label12.Caption = Format(Now, "hh:mm:ss")
DoEvents
End Sub

Private Sub Timer1_Timer()
If Text12.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text12.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text41.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub

Private Sub Timer10_Timer()
DoEvents
If Text21.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text21.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text32.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub

Private Sub Timer11_Timer()
DoEvents
If Text22.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer12_Timer()
DoEvents
If Text23.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer13_Timer()
DoEvents
If Text24.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer14_Timer()
DoEvents
If Text25.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer15_Timer()
DoEvents
If Text26.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer16_Timer()
DoEvents
If Text27.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer17_Timer()
DoEvents
If Text28.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer18_Timer()
DoEvents
If Text29.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer19_Timer()
DoEvents
If Text30.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer2_Timer()
DoEvents
If Text13.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text13.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text40.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub

Private Sub Timer20_Timer()
DoEvents
If Text31.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + ExecuteCommandLineOutput("taskkill /f /im wemeetapp.exe") & vbCrLf
DoEvents
Sleep 1000 'ms
CreateObject("WScript.Shell").Run """C:\Program Files (x86)\Tencent\WeMeet\wemeetapp.exe""", 0
DoEvents
End If
DoEvents
End Sub

Private Sub Timer3_Timer()
DoEvents
If Text14.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text14.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text39.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub

Private Sub Timer4_Timer()
DoEvents
If Text15.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text15.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text38.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub

Private Sub Timer5_Timer()
DoEvents
If Text16.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text16.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text37.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents

DoEvents
End If
DoEvents
End Sub

Private Sub Timer6_Timer()
DoEvents
If Text17.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text17.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text36.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub

Private Sub Timer7_Timer()
DoEvents
If Text18.Text = Label12.Caption Then
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text35.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub

Private Sub Timer8_Timer()
DoEvents
If Text19.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text19.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text34.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub

Private Sub Timer9_Timer()
DoEvents
If Text20.Text = Label12.Caption Then
DoEvents
Text1.Text = Text1.Text + "---" + Text20.Text + "---" & vbCrLf
DoEvents
ExecuteCommandLineOutput (Text3.Text + " wemeet://page/inmeeting?meeting_code=" + Text33.Text)
DoEvents
Text1.Text = Text1.Text + "执行完毕!!!" & vbCrLf
DoEvents
Sleep 1000 'ms
DoEvents
End If
DoEvents
End Sub
