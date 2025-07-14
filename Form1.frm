VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Tetris VB6 2025"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   7050
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4920
      Top             =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton box 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   450
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntaje:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   1335
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

Dim Score As Integer

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
Private m_Grid(GRID_WIDTH, GRID_HEIGHT) As Boolean

' Inicializar el juego
Private Sub InitializeGame()
    ' Configurar el tamaño del Frame
    Frame1.Width = GRID_WIDTH * BOX_SIZE
    Frame1.Height = GRID_HEIGHT * BOX_SIZE
    
    Score = 0
    
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


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyW: RotatePiece
        Case vbKeyA: MoveLeft
        Case vbKeyD: MoveRight
        Case vbKeyS: MovePiece 0, BOX_SIZE
        Case vbKeySpace: InstantDrop
    End Select
End Sub

Private Sub Form_Load()
    ' Inicializar el generador de números aleatorios
    Randomize
    
    ' Inicializar el juego
    InitializeGame
    
    ' Asegurarse de que el formulario pueda recibir eventos de teclado
    KeyPreview = True
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
    ' Mover la pieza hacia abajo
    MovePiece 0, BOX_SIZE
End Sub

Private Sub MoveLeft()
    MovePiece -BOX_SIZE, 0
End Sub

Private Sub MoveRight()
    MovePiece BOX_SIZE, 0
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

Private Function AlignToGrid(value As Single) As Integer
    ' Redondea un valor a la coordenada de la grilla más cercana
    ' Se usa Int(x + 0.5) para un redondeo aritmético estándar y evitar
    ' el comportamiento de CInt que redondea al par más cercano en .5
    AlignToGrid = Int(value / BOX_SIZE + 0.5) * BOX_SIZE
End Function



Private Function CheckCollision(offsetX As Integer, offsetY As Integer) As Boolean
    Dim activeBlock As CommandButton
    Dim gridX As Integer, gridY As Integer

    For Each activeBlock In m_ActiveBlocks
        gridX = (activeBlock.Left + offsetX) / BOX_SIZE
        gridY = (activeBlock.Top + offsetY) / BOX_SIZE

        ' 1. Verificar colisión con bordes
        If gridX < 0 Or gridX >= GRID_WIDTH Or gridY >= GRID_HEIGHT Then
            CheckCollision = True
            Exit Function
        End If

        ' 2. Verificar colisión con bloques aterrizados en la grilla
        If m_Grid(gridX, gridY) Then
            CheckCollision = True
            Exit Function
        End If
    Next activeBlock

    CheckCollision = False
End Function

Private Sub LandPiece()
    Dim block As CommandButton
    Dim gridX As Integer, gridY As Integer
    
    ' Marcar la posición de la pieza en la grilla y transferir los bloques
    For Each block In m_ActiveBlocks
        gridX = block.Left / BOX_SIZE
        gridY = block.Top / BOX_SIZE
        If gridX >= 0 And gridX < GRID_WIDTH And gridY >= 0 And gridY < GRID_HEIGHT Then
            m_Grid(gridX, gridY) = True
        End If
        m_LandedBlocks.Add block
    Next
    
    ' Limpiar líneas completas
    ClearCompletedLines
    
    ' Generar nueva pieza
    ShowRandomPiece
End Sub

Private Sub ClearCompletedLines()
    Dim y As Integer, x As Integer, x2 As Integer, y2 As Integer
    Dim fullLine As Boolean
    
    For y = GRID_HEIGHT - 1 To 0 Step -1
        fullLine = True
        For x = 0 To GRID_WIDTH - 1
            If Not m_Grid(x, y) Then
                fullLine = False
                Exit For
            End If
        Next x
        
        If fullLine Then
            ' Eliminar los botones de la línea
            For i = m_LandedBlocks.Count To 1 Step -1
                If CInt(m_LandedBlocks(i).Top / BOX_SIZE) = y Then
                    Unload m_LandedBlocks(i)
                    m_LandedBlocks.Remove i
                End If
            Next i
            
            ' Bajar las líneas superiores en la grilla
            For y2 = y To 1 Step -1
                For x2 = 0 To GRID_WIDTH - 1
                    m_Grid(x2, y2) = m_Grid(x2, y2 - 1)
                Next x2
            Next y2
            
            ' Limpiar la fila superior de la grilla
            For x2 = 0 To GRID_WIDTH - 1
                m_Grid(x2, 0) = False
            Next x2
            
            ' Bajar los botones de las líneas superiores
            For Each block In m_LandedBlocks
                If (block.Top / BOX_SIZE) < y Then
                    block.Top = block.Top + BOX_SIZE
                End If
            Next block
            
            ' Repetir la comprobación para la misma línea y (que ahora es nueva)
            y = y + 1
            
            ' Mostramos puntaje
            Score = Score + 1
            Label2.Caption = Score
        End If
    Next y
End Sub

Private Sub InstantDrop()
    ' Mueve la pieza hacia abajo hasta que haya colisión
    Do While Not CheckCollision(0, BOX_SIZE)
        MovePiece 0, BOX_SIZE
        ' Pequeña pausa para la animación
        DoEvents
    Loop
    ' Asegurarse de que la pieza se coloque correctamente
    LandPiece
End Sub

Private Sub MovePiece(offsetX As Integer, offsetY As Integer)
    ' Verifica si la nueva posición es válida
    If Not CheckCollision(offsetX, offsetY) Then
        ' Si es válida, mueve la pieza
        Dim block As CommandButton
        For Each block In m_ActiveBlocks
            block.Left = block.Left + offsetX
            block.Top = block.Top + offsetY
        Next block
    Else
        ' Si no es válida y el movimiento era hacia abajo, bloquea la pieza
        If offsetY > 0 Then
            LandPiece
        End If
    End If
End Sub


