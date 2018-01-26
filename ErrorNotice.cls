VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pErrorExists As Boolean
Private pDescription As String
Private pCusipProblemExists As Boolean
Private pCusipProblemNumber As Integer
Private pTickerProblemExists As Boolean
Private pTickerProblemNumber As Integer

Private Sub Class_Initialize()
    ErrorExists = False
    CusipProblemExists = False
    Description = "No Error"
    pCusipProblemNumber = 0
End Sub

Public Property Get ErrorExists() As Boolean
    ErrorExists = pErrorExists
End Property

Public Property Let ErrorExists(ByVal booNewValue As Boolean)
    pErrorExists = booNewValue
End Property


Public Property Get Description() As String
    Description = pDescription
End Property

Public Property Let Description(ByVal strValue As String)
    pDescription = strValue
End Property



Public Property Get CusipProblemExists() As Boolean
    CusipProblemExists = pCusipProblemExists
End Property

Public Property Let CusipProblemExists(ByVal vNewValue As Boolean)
    pCusipProblemExists = vNewValue
End Property

Public Property Get CusipProblemNumber() As Integer
    CusipProblemNumber = pCusipProblemNumber
End Property

Public Property Let CusipProblemNumber(ByVal intNewValue As Integer)
    pCusipProblemNumber = intNewValue
End Property


Public Property Get TickerProblemExists() As Boolean
    TickerProblemExists = pTickerProblemExists
End Property

Public Property Let TickerProblemExists(ByVal vNewValue As Boolean)
    pTickerProblemExists = vNewValue
End Property

Public Property Get TickerProblemNumber() As Integer
    TickerProblemNumber = pTickerProblemNumber
End Property

Public Property Let TickerProblemNumber(ByVal intNewValue As Integer)
    pTickerProblemNumber = intNewValue
End Property
