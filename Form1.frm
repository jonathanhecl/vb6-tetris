VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tetris VB6 2025"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   7920
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Acelerar (&S)"
      Height          =   735
      Left            =   5760
      TabIndex        =   5
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Derecha (&D)"
      Height          =   735
      Left            =   6600
      TabIndex        =   4
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Izquierda (&A)"
      Height          =   735
      Left            =   4920
      TabIndex        =   3
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rotar (&W)"
      Height          =   735
      Left            =   5760
      TabIndex        =   2
      Top             =   8280
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4920
      Top             =   960
   End
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
Private m_ActiveBlocks As Collection
Private m_LandedBlocks As Collection
Private m_CurrentPieceType As String
Private m_CurrentRotation As Integer ' 0: 0°, 1: 90°, 2: 180°, 3: 270°

' Inicializar el juego
Private Sub InitializeGame()
    ' Configurar el tamaño del Frame
    Frame1.Width = GRID_WIDTH * BOX_SIZE
    Frame1.Height = GRID_HEIGHT * BOX_SIZE
    
    ' Inicializar piezas
    InitializePieceTypes
    
    ' Inicializar colecciones
    Set m_ActiveBlocks = New Collection
    Set m_LandedBlocks = New Collection

    ' Mostrar la primera pieza
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
    
    ' Calcular la posición centrada en la grilla
    Dim gridWidthInBlocks As Integer
    Dim startCol As Integer
    gridWidthInBlocks = Frame1.Width \ BOX_SIZE
    startCol = (gridWidthInBlocks - m_PieceTypes(pieceIndex).Width) \ 2
    startX = startCol * BOX_SIZE
    
    ' Limpiar bloques activos anteriores
    Set m_ActiveBlocks = New Collection

    ' Inicializar rotación
    m_CurrentRotation = 0
    m_CurrentPieceType = m_PieceTypes(pieceIndex).Name
    
    ' Crear la pieza
    CreatePiece startX, 0, m_CurrentPieceType, m_PieceTypes(pieceIndex).Color
End Sub

' Obtener indice de pieza aleatoria
Private Function GetRandomPieceIndex() As Integer
    GetRandomPieceIndex = Int(Rnd * 7) ' 7 tipos de piezas (0-6)
End Function

Private Sub Command1_Click()
    Call RotatePiece
End Sub

Private Sub Command2_Click()
    Call MoveLeft
End Sub

Private Sub Command3_Click()
    Call MoveRight
End Sub

Private Sub Command4_Click()
    If Timer1.Interval >= 200 Then
        Timer1.Interval = Timer1.Interval - 100
    End If
End Sub

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
                m_ActiveBlocks.Add btn
            Next i
            
        Case "O" ' O shape (square)
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "1", startX + BOX_SIZE, startY, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "2", startX, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "3", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            
        Case "J" ' J shape
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "1", startX, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "2", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "3", startX + (2 * BOX_SIZE), startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            
        Case "L" ' L shape
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "1", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "2", startX + (2 * BOX_SIZE), startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "3", startX + (2 * BOX_SIZE), startY, pieceColor)
            m_ActiveBlocks.Add btn
            
        Case "S" ' S shape
            Set btn = CreateButton("btn" & pieceType & "0", startX + BOX_SIZE, startY, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "1", startX + (2 * BOX_SIZE), startY, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "2", startX, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "3", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            
        Case "Z" ' Z shape
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "1", startX + BOX_SIZE, startY, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "2", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "3", startX + (2 * BOX_SIZE), startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            
        Case "T" ' T shape
            Set btn = CreateButton("btn" & pieceType & "0", startX, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "1", startX + BOX_SIZE, startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "2", startX + (2 * BOX_SIZE), startY + BOX_SIZE, pieceColor)
            m_ActiveBlocks.Add btn
            Set btn = CreateButton("btn" & pieceType & "3", startX + BOX_SIZE, startY, pieceColor)
            m_ActiveBlocks.Add btn
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

Private Sub Timer1_Timer()
    Dim block As CommandButton

    If CanMoveDown() Then
        ' Mover cada bloque de la pieza activa hacia abajo
        For Each block In m_ActiveBlocks
            block.Top = block.Top + BOX_SIZE
        Next block
    Else
        ' La pieza ha aterrizado, transferirla a los bloques "aterrizados"
        For Each block In m_ActiveBlocks
            m_LandedBlocks.Add block
        Next block
        
        ' Reiniciamos el tiempo
        Timer1.Interval = 500
        
        ' Generar una nueva pieza
        ShowRandomPiece
    End If
End Sub

Private Sub MoveLeft()
    ' Solo verificar colisión con el borde izquierdo, no con bloques aterrizados
    Dim canMove As Boolean
    canMove = True
    
    ' Verificar colisión con el borde izquierdo
    Dim block As CommandButton
    For Each block In m_ActiveBlocks
        If block.Left - BOX_SIZE < 0 Then
            canMove = False
            Exit For
        End If
    Next block
    
    ' Si no hay colisión con el borde, mover la pieza
    If canMove Then
        For Each block In m_ActiveBlocks
            block.Left = block.Left - BOX_SIZE
        Next block
    End If
End Sub

Private Sub MoveRight()
    ' Solo verificar colisión con el borde derecho, no con bloques aterrizados
    Dim canMove As Boolean
    canMove = True
    
    ' Verificar colisión con el borde derecho
    Dim block As CommandButton
    For Each block In m_ActiveBlocks
        If block.Left + BOX_SIZE >= Frame1.Width Then
            canMove = False
            Exit For
        End If
    Next block
    
    ' Si no hay colisión con el borde, mover la pieza
    If canMove Then
        For Each block In m_ActiveBlocks
            block.Left = block.Left + BOX_SIZE
        Next block
    End If
End Sub

Private Function CanRotate(blocks() As Integer) As Boolean
    ' Verificar si la rotación es válida (sin colisiones)
    Dim i As Integer
    Dim testX As Integer, testY As Integer
    
    ' Verificar colisiones con bordes y otras piezas
    For i = 0 To UBound(blocks) Step 2
        testX = blocks(i)
        testY = blocks(i + 1)
        
        ' Verificar colisión con bordes
        If testX < 0 Or testX >= Frame1.Width Or testY < 0 Or testY >= Frame1.Height Then
            CanRotate = False
            Exit Function
        End If
        
        ' Verificar colisión con bloques aterrizados
        Dim landedBlock As CommandButton
        For Each landedBlock In m_LandedBlocks
            If testX = landedBlock.Left And testY = landedBlock.Top Then
                CanRotate = False
                Exit Function
            End If
        Next landedBlock
    Next i
    
    CanRotate = True
End Function

Private Sub RotatePiece()
    ' No rotar la pieza O (cuadrado)
    If m_CurrentPieceType = "O" Then Exit Sub
    
    ' Calcular el centro de rotación (usando el segundo bloque como pivote)
    If m_ActiveBlocks.Count < 2 Then Exit Sub
    
    Dim centerX As Integer, centerY As Integer
    centerX = m_ActiveBlocks(2).Left + BOX_SIZE \ 2
    centerY = m_ActiveBlocks(2).Top + BOX_SIZE \ 2
    
    ' Calcular nuevas posiciones después de la rotación
    Dim newPositions() As Integer
    ReDim newPositions((m_ActiveBlocks.Count * 2) - 1)
    
    Dim i As Integer, j As Integer
    Dim relX As Integer, relY As Integer
    Dim newX As Integer, newY As Integer
    
    j = 0
    For Each block In m_ActiveBlocks
        ' Calcular posición relativa al centro
        relX = block.Left + BOX_SIZE \ 2 - centerX
        relY = block.Top + BOX_SIZE \ 2 - centerY
        
        ' Aplicar rotación 90° en sentido horario: (x,y) -> (y,-x)
        newX = centerX + relY - BOX_SIZE \ 2
        newY = centerY - relX - BOX_SIZE \ 2
        
        ' Asegurar que las coordenadas estén alineadas con la grilla
        newX = (newX \ BOX_SIZE) * BOX_SIZE
        newY = (newY \ BOX_SIZE) * BOX_SIZE
        
        ' Guardar la nueva posición
        newPositions(j) = newX
        newPositions(j + 1) = newY
        j = j + 2
    Next block
    
    ' Verificar si la rotación es válida
    If CanRotate(newPositions) Then
        ' Actualizar las posiciones de los bloques
        j = 0
        For Each block In m_ActiveBlocks
            block.Left = newPositions(j)
            block.Top = newPositions(j + 1)
            j = j + 2
        Next block
        
        ' Actualizar el estado de rotación
        m_CurrentRotation = (m_CurrentRotation + 1) Mod 4
    End If
End Sub

Private Function GetPieceColor(pieceType As String) As Long
    Dim i As Integer
    For i = 0 To UBound(m_PieceTypes)
        If m_PieceTypes(i).Name = pieceType Then
            GetPieceColor = m_PieceTypes(i).Color
            Exit Function
        End If
    Next i
    GetPieceColor = vbBlack ' Color por defecto
End Function

Private Function CanMoveDown() As Boolean
    Dim activeBlock As CommandButton
    Dim landedBlock As CommandButton
    CanMoveDown = True ' Asumir que se puede mover

    For Each activeBlock In m_ActiveBlocks
        ' 1. Comprobar si el bloque ha llegado al fondo del area de juego
        If activeBlock.Top + BOX_SIZE >= Frame1.Height Then
            CanMoveDown = False
            Exit Function
        End If
        
        ' 2. Comprobar si choca con un bloque ya aterrizado
        For Each landedBlock In m_LandedBlocks
            If activeBlock.Left = landedBlock.Left And activeBlock.Top + BOX_SIZE = landedBlock.Top Then
                CanMoveDown = False
                Exit Function
            End If
        Next landedBlock
    Next activeBlock
End Function

