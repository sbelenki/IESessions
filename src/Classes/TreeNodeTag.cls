VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeNodeTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TreeNodeType
    Root
    Session
    Uri
End Enum

Private m_nNodeType As TreeNodeType
Private m_nKey As Integer

Property Get NodeType( _
    ) As TreeNodeType
    
    NodeType = m_nNodeType
End Property

Property Let NodeType( _
    xNewType As TreeNodeType)
    
    m_nNodeType = xNewType
End Property

Property Get Key( _
    ) As Integer
    
    Key = m_nKey
End Property

Property Let Key( _
    xNewKey As Integer)
    
    m_nKey = xNewKey
End Property

