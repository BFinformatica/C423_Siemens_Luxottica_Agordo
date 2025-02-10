VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TagAnalogica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarNome As String 'Copia locale.
Private mvarValore As Variant 'Copia locale.
Public Property Let Valore(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.Valore = 5
    mvarValore = vData
End Property


Public Property Set Valore(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.Valore = Form1
    Set mvarValore = vData
End Property


Public Property Get Valore() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Valore
    If IsObject(mvarValore) Then
        Set Valore = mvarValore
    Else
        Valore = mvarValore
    End If
End Property

Public Property Let Nome(ByVal vData As Variant)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.Nome = 5
    mvarNome = vData
End Property


Public Property Get Nome() As Variant
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Nome
    Nome = mvarNome
End Property



