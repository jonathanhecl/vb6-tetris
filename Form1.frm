VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tetris VB6 2025"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton box 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Constantes globales
Const BOX_SIZE As Integer = 400
Const GRID_WIDTH As Integer = 10
Const GRID_HEIGHT As Integer = 24

' Tipos de piezas disponibles
Private Type pieceType
    Name As String
    Color As Long
    Width As Integer
    Height As Integer
End Type

Private m_PieceTypes() As pieceType

' Inicializar el juego
Private Sub InitializeGame()
    ' Configurar el tamaÃ±o del Frame
    Frame1.Width = GRID_WIDTH * BOX_SIZE
    Frame1.Height = GRID_HEIGHT * BOX_SIZE
    
    ' Inicializar piezas
    InitializePieceTypes
    
    ' Mostrar una pieza de prueba
    ShowRandomPiece
End Sub

' Inicializar los tipos de piezas
Private Sub InitializePieceTypes()
    ReDim m_PieceTypes(6) As pieceType
    
    With m_PieceTypes(0) ' I (linea)
        .Name = "I": .Color = vbCyan: .Width = 4: .Height = 1
    End With
    With m_PieceTypes(1) ' O (cuadrado)
        .Name = "O": .Color = vbYellow: .Width = 2: .Height = 2
    End With
    With m_PieceTypes(2) ' J
        .Name = "J": .Color = vbBlue: .Width = 2: .Height = 3
    End With
    With m_PieceTypes(3) ' L (naranja)
        .Name = "L": .Color = &HFF8000: .Width = 2: .Height = 3
    End With
    With m_PieceTypes(4) ' S
        .Name = "S": .Color = vbGreen: .Width = 3: .Height = 2
    End With
    With m_PieceTypes(5) ' Z
        .Name = "Z": .Color = vbRed: .Width = 3: .Height = 2
    End With
    With m_PieceTypes(6) ' T (morado)
        .Name = "T": .Color = &HFF00FF: .Width = 3: .Height = 2
    End With
End Sub

' Mostrar una pieza aleatoria
Private Sub ShowRandomPiece()
    Dim pieceIndex As Integer
    Dim startX As Integer
    
    ' Seleccionar una pieza al azar
    pieceIndex = GetRandomPieceIndex()
    
    ' Calcular la posición centrada
    startX = (Frame1.Width - (m_PieceTypes(pieceIndex).Width * BOX_SIZE)) \ 2
    
    ' Crear la pieza
    CreatePiece startX, 0, m_PieceTypes(pieceIndex).Name, m_PieceTypes(pieceIndex).Color
End Sub

' Obtener indice de pieza aleatoria
Private Function GetRandomPieceIndex() As Integer
    GetRandomPieceIndex = Int(Rnd * 7) ' 7 tipos de piezas (0-6)
End Function

Private Sub Form_Load()
    ' Inicializar el generador de números aleatorios
    Randomize
    
    ' Inicializar el juego
    InitializeGame
End Sub

' Crea una pieza en la posición especificada
Private Sub CreatePiece(startX As Integer, startY As Integer, pieceType As String, pieceColor As Long)
    Dim i As Integer
    Dim btn As CommandButton
    
    Select Case pieceType
        Case "I" ' I shape (line)
            For i = 0 To 3
                Set btn = CreateButton("btn" & pieceType & i, startX + (i * BOX_SIZE), startY, pieceColor)
            Next i
            
        Case "O" ' O shape (square)
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "1", startX + BOX_SIZE, startY, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "2", startX, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "3", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            
        Case "J" ' J shape
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "1", startX, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "2", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "3", startX + (2 * BOX_SIZE), startY + BOX_SIZE, pieceColor)
            
        Case "L" ' L shape
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "1", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "2", startX + (2 * BOX_SIZE), startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "3", startX + (2 * BOX_SIZE), startY, pieceColor)
            
        Case "S" ' S shape
            Set btn = CreateButton("btn" & pieceType & "0", startX + BOX_SIZE, startY, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "1", startX + (2 * BOX_SIZE), startY, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "2", startX, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "3", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            
        Case "Z" ' Z shape
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "1", startX + BOX_SIZE, startY, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "2", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "3", startX + (2 * BOX_SIZE), startY + BOX_SIZE, pieceColor)
            
        Case "T" ' T shape
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "1", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "2", startX + (2 * BOX_SIZE), startY + BOX_SIZE, pieceColor)
            Set btn = CreateButton("btn" & pieceType & "3", startX + BOX_SIZE, startY, pieceColor)
    End Select
End Sub

Private Function CreateButton(btnName As String, x As Integer, y As Integer, btnColor As Long) As CommandButton
    Static buttonCount As Long
    buttonCount = buttonCount + 1
    
    ' Crear una nueva instancia del control array box
    Load box(buttonCount)
    
    With box(buttonCount)
        .BackColor = btnColor
        .Width = BOX_SIZE - 2  ' Pequeño espacio entre bloques
        .Height = BOX_SIZE - 2  ' Pequeño espacio entre bloques
        .Left = x + 1
        .Top = y + 1
        .Visible = True
        Set .Container = Frame1  ' Asegurar que el botón está dentro del Frame1
    End With
    
    Set CreateButton = box(buttonCount)
End Function
