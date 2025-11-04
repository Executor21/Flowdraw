/*
Script: Flowdraw
Συγγραφέας: Tasos
Έτος: 2025
MIT License
Copyright (c) 2025 Tasos
*/

#Requires AutoHotkey v2.0
#SingleInstance Force

; === CONFIGURATION ===
class Config {
    static CANVAS_WIDTH := 1100   ; Μειώθηκε από 1200
    static CANVAS_HEIGHT := 680   ; ΑΥΞΗΘΗΚΕ από 600
    static GRID_SIZE := 20
    static GRID_MAJOR := 100
    static TIMER_INTERVAL := 50
    static MAX_TEXT_LENGTH := 200
    static SHAPE_DEFAULT_WIDTH := 120
    static SHAPE_DEFAULT_HEIGHT := 60
}

; === GLOBAL VARIABLES ===
global Shapes := Map()
global Connections := Array()
global SelectedShape := ""
global IsDrawingConnection := false
global StartShape := ""
global CurrentShapeType := ""
global ShapeCount := 0
global IsDragging := false
global DragOffsetX := 0, DragOffsetY := 0
global LastMouseX := 0, LastMouseY := 0
global SelectedConnection := ""
global RedrawScheduled := false

; Undo/Redo System
global UndoStack := Array()
global RedoStack := Array()

; GDI+ Globals
global gpToken := 0
global hBitmap := 0
global hdcMem := 0
global G := 0
global CLSID_PNG := ""

; Canvas dimensions (dynamic)
global CanvasWidth := Config.CANVAS_WIDTH
global CanvasHeight := Config.CANVAS_HEIGHT

; GUI Controls
global MyGui := ""
global Canvas := ""
global StatusText := ""
global CurrentSelectionText := ""
global StatsText := ""

; === INITIALIZATION ===
TraySetIcon("Shell32.dll", 44)

MyGui := Gui()
MyGui.Title := "Flowdraw"
MyGui.Opt("-Resize +MaximizeBox +MinimizeBox")
MyGui.OnEvent("Close", GuiClose)
MyGui.OnEvent("ContextMenu", GuiContextMenu)
MyGui.SetFont("s9", "Arial")

MyGui.Add("Text", "x10 y10 w1000 h30", "Click+Drag to move shapes | Right-click canvas to connect | Double-click connections to disconnect")

Canvas := MyGui.Add("Text", "x10 y50 w" CanvasWidth " h" CanvasHeight " +Border +0x1000", "")
Canvas.SetFont("s9", "Arial")
Canvas.OnEvent("Click", CanvasClick)
Canvas.OnEvent("DoubleClick", CanvasDoubleClick)

StatusText := MyGui.Add("Text", "x10 y740 w1100 h20", "Initializing...")

; Sidebar - Single column with proper spacing
MyGui.Add("Text", "x1120 y10 w240 h20", "📐 Shape Palette:")
MyGui.Add("Button", "x1120 y35 w240 h32", "🟦 Process").OnEvent("Click", (*) => SelectShapeType("rectangle"))
MyGui.Add("Button", "x1120 y72 w240 h32", "🔷 Decision").OnEvent("Click", (*) => SelectShapeType("diamond")) 
MyGui.Add("Button", "x1120 y109 w240 h32", "⭕ Start/End").OnEvent("Click", (*) => SelectShapeType("ellipse"))
MyGui.Add("Button", "x1120 y146 w240 h32", "📊 Data").OnEvent("Click", (*) => SelectShapeType("data"))
MyGui.Add("Button", "x1120 y183 w240 h32", "📄 Document").OnEvent("Click", (*) => SelectShapeType("document"))

MyGui.Add("Text", "x1120 y230 w240 h20", "🛠️ Tools:")
MyGui.Add("Button", "x1120 y255 w240 h32", "🔗 Connect").OnEvent("Click", ConnectTool)
MyGui.Add("Button", "x1120 y292 w240 h32", "✏️ Edit").OnEvent("Click", EditText)
MyGui.Add("Button", "x1120 y329 w240 h32", "➖ Disconnect").OnEvent("Click", DisconnectTool)
MyGui.Add("Button", "x1120 y366 w240 h32", "🗑️ Delete").OnEvent("Click", DeleteSelected)
MyGui.Add("Button", "x1120 y403 w240 h32", "🧹 Clear All").OnEvent("Click", ClearAll)

MyGui.Add("Text", "x1120 y450 w240 h20", "💾 File Operations:")
MyGui.Add("Button", "x1120 y475 w240 h32", "💾 Save Project").OnEvent("Click", SaveProjectINI)
MyGui.Add("Button", "x1120 y512 w240 h32", "📂 Load Project").OnEvent("Click", LoadProjectINI)
MyGui.Add("Button", "x1120 y549 w240 h32", "📷 Export PNG").OnEvent("Click", ExportToPNG)

MyGui.Add("Text", "x1120 y596 w120 h20", "📋 Current Selection:")
CurrentSelectionText := MyGui.Add("Text", "x1120 y621 w120 h30", "None")

MyGui.Add("Text", "x1250 y596 w120 h20", "📊 Statistics:")
StatsText := MyGui.Add("Text", "x1250 y621 w120 h30", "Shapes: 0`nConnections: 0")

MyGui.Add("Text", "x1120 y661 w240 h20", "🔄 History:")
MyGui.Add("Button", "x1120 y685 w115 h26", "↶ Undo").OnEvent("Click", UndoAction)
MyGui.Add("Button", "x1245 y685 w115 h26", "↷ Redo").OnEvent("Click", RedoAction)

MyGui.Add("Button", "x1120 y717 w240 h30", "💡 Tips").OnEvent("Click", ShowTips)

; Initialize GDI+
if !InitializeGDIPlus() {
    MsgBox("Failed to initialize GDI+!`n`nThe application cannot start.`n`nPlease ensure GDI+ is available on your system.", 
           "Critical Error", "Icon! 16")
    ExitApp
}

InitEncoderCLSID()

if !InitGDI() {
    MsgBox("Failed to initialize drawing canvas!`n`nThe application cannot start.", 
           "Critical Error", "Icon! 16")
    ExitApp
}

MyGui.Show("w1366 h768")  ; Μείωσα λίγο το ύψος τώρα που αφαιρέθηκαν τα tips
SetTimer(CheckMouse, Config.TIMER_INTERVAL)
ScheduleRedraw()
UpdateStatus("Ready - Edition v1.0 loaded successfully!")

; === TIPS WINDOW ===
ShowTips(*) {
    TipsGui := Gui()
    TipsGui.Title := "Flowdraw Tips & Shortcuts"
    TipsGui.OnEvent("Close", (*) => TipsGui.Destroy())
    TipsGui.SetFont("s10", "Arial")
    
    ; Main instructions
    TipsGui.Add("Text", "x10 y10 w500 h30", "🎯 Basic Operations:")
    TipsGui.Add("Text", "x20 y40 w480 h80", 
        "• Click on canvas to place selected shape`n" .
        "• Drag shapes to move them`n" .
        "• Right-click two shapes to connect them`n" .
        "• Double-click connections to remove them`n" .
        "• Click shapes to select, then use Edit/Delete")
    
    ; Selection & Editing
    TipsGui.Add("Text", "x10 y130 w500 h30", "✏️ Selection & Editing:")
    TipsGui.Add("Text", "x20 y160 w480 h60", 
        "• Click any shape to select it`n" .
        "• Use Edit button to change text`n" .
        "• Use Delete button to remove selected`n" .
        "• ESC to cancel current operation")
    
    ; Keyboard Shortcuts
    TipsGui.Add("Text", "x10 y230 w500 h30", "⌨️ Keyboard Shortcuts:")
    TipsGui.Add("Text", "x20 y260 w480 h160", 
        "• Ctrl+1: Process shape`n" .
        "• Ctrl+2: Start/End shape`n" .
        "• Ctrl+3: Decision shape`n" .
        "• Ctrl+C: Connect mode`n" .
        "• Ctrl+D: Disconnect mode`n" .
        "• Ctrl+Z: Undo`n" .
        "• Ctrl+Y: Redo`n" .
        "• Delete: Remove selected`n" .
        "• Ctrl+S: Save project`n" .
        "• Ctrl+O: Load project`n" .
        "• Ctrl+E: Export to PNG")
    
    ; File Operations
    TipsGui.Add("Text", "x10 y430 w500 h30", "💾 File Operations:")
    TipsGui.Add("Text", "x20 y460 w480 h60", 
        "• Save Project: Saves as INI file`n" .
        "• Load Project: Loads from INI file`n" .
        "• Export PNG: Creates image of your diagram")
    
    ; Close button
    TipsGui.Add("Button", "x200 y530 w100 h35", "Close").OnEvent("Click", (*) => TipsGui.Destroy())
    
    TipsGui.Show("w360 h580")
}

; === UNDO/REDO SYSTEM ===

AddUndoStep(action, data) {
    global UndoStack, RedoStack
    
    ; Keep only last 50 actions to prevent memory issues
    if UndoStack.Length >= 50 {
        UndoStack.RemoveAt(1)
    }
    
    UndoStack.Push(Map("action", action, "data", data, "timestamp", A_TickCount))
    RedoStack := Array() ; Clear redo stack when new action is performed
}

UndoAction(*) {
    global UndoStack, RedoStack, Shapes, Connections, SelectedShape, SelectedConnection
    
    if UndoStack.Length = 0 {
        UpdateStatus("Nothing to undo")
        return
    }
    
    lastAction := UndoStack.Pop()
    RedoStack.Push(lastAction)
    
    switch lastAction["action"] {
        case "add_shape":
            shapeId := lastAction["data"]["id"]
            if Shapes.Has(shapeId) {
                Shapes.Delete(shapeId)
                if SelectedShape = shapeId {
                    SelectedShape := ""
                }
                RemoveConnections(shapeId)
            }
            UpdateStatus("Undo: Added shape")
            
        case "delete_shape":
            Shapes[lastAction["data"]["id"]] := lastAction["data"]["shape"]
            UpdateStatus("Undo: Deleted shape")
            
        case "move_shape":
            shapeId := lastAction["data"]["id"]
            if Shapes.Has(shapeId) {
                Shapes[shapeId]["x"] := lastAction["data"]["oldX"]
                Shapes[shapeId]["y"] := lastAction["data"]["oldY"]
            }
            UpdateStatus("Undo: Moved shape")
            
        case "edit_text":
            shapeId := lastAction["data"]["id"]
            if Shapes.Has(shapeId) {
                Shapes[shapeId]["text"] := lastAction["data"]["oldText"]
            }
            UpdateStatus("Undo: Edited text")
            
        case "add_connection":
            if lastAction["data"].Has("index") && lastAction["data"]["index"] <= Connections.Length {
                Connections.RemoveAt(lastAction["data"]["index"])
            }
            UpdateStatus("Undo: Added connection")
            
        case "delete_connection":
            Connections.Push(lastAction["data"]["connection"])
            UpdateStatus("Undo: Deleted connection")
            
        case "clear_all":
            Shapes := lastAction["data"]["shapes"].Clone()
            Connections := lastAction["data"]["connections"].Clone()
            UpdateStatus("Undo: Cleared all")
    }
    
    StatsText.Text := "Shapes: " Shapes.Count "`nConnections: " Connections.Length
    ScheduleRedraw()
}

RedoAction(*) {
    global UndoStack, RedoStack, Shapes, Connections
    
    if RedoStack.Length = 0 {
        UpdateStatus("Nothing to redo")
        return
    }
    
    lastAction := RedoStack.Pop()
    UndoStack.Push(lastAction)
    
    switch lastAction["action"] {
        case "add_shape":
            Shapes[lastAction["data"]["id"]] := lastAction["data"]["shape"]
            UpdateStatus("Redo: Added shape")
            
        case "delete_shape":
            shapeId := lastAction["data"]["id"]
            if Shapes.Has(shapeId) {
                Shapes.Delete(shapeId)
                RemoveConnections(shapeId)
            }
            UpdateStatus("Redo: Deleted shape")
            
        case "move_shape":
            shapeId := lastAction["data"]["id"]
            if Shapes.Has(shapeId) {
                Shapes[shapeId]["x"] := lastAction["data"]["newX"]
                Shapes[shapeId]["y"] := lastAction["data"]["newY"]
            }
            UpdateStatus("Redo: Moved shape")
            
        case "edit_text":
            shapeId := lastAction["data"]["id"]
            if Shapes.Has(shapeId) {
                Shapes[shapeId]["text"] := lastAction["data"]["newText"]
            }
            UpdateStatus("Redo: Edited text")
            
        case "add_connection":
            Connections.Push(lastAction["data"]["connection"])
            UpdateStatus("Redo: Added connection")
            
        case "delete_connection":
            if lastAction["data"].Has("index") && lastAction["data"]["index"] <= Connections.Length {
                Connections.RemoveAt(lastAction["data"]["index"])
            }
            UpdateStatus("Redo: Deleted connection")
            
        case "clear_all":
            Shapes := Map()
            Connections := Array()
            UpdateStatus("Redo: Cleared all")
    }
    
    StatsText.Text := "Shapes: " Shapes.Count "`nConnections: " Connections.Length
    ScheduleRedraw()
}

; === CORE FUNCTIONS ===

InitializeGDIPlus() {
    global gpToken
    
    if !DllCall("GetModuleHandle", "str", "gdiplus", "ptr") {
        if !DllCall("LoadLibrary", "str", "gdiplus", "ptr") {
            LogError("Failed to load gdiplus.dll")
            return false
        }
    }
    
    si := Buffer(24, 0)
    NumPut("uint", 1, si, 0)
    
    result := DllCall("gdiplus\GdiplusStartup", "ptr*", &gpToken, "ptr", si.Ptr, "ptr", 0, "int")
    
    if (result != 0) {
        LogError("GdiplusStartup failed with error code: " result)
        return false
    }
    
    return true
}

InitEncoderCLSID() {
    global CLSID_PNG
    
    CLSID_PNG := Buffer(16)
    NumPut("uint", 0x557CF406, CLSID_PNG, 0)
    NumPut("ushort", 0x1A04, CLSID_PNG, 4)
    NumPut("ushort", 0x11D3, CLSID_PNG, 6)
    NumPut("uchar", 0x9A, CLSID_PNG, 8)
    NumPut("uchar", 0x73, CLSID_PNG, 9)
    NumPut("uchar", 0x00, CLSID_PNG, 10)
    NumPut("uchar", 0x00, CLSID_PNG, 11)
    NumPut("uchar", 0xF8, CLSID_PNG, 12)
    NumPut("uchar", 0x1E, CLSID_PNG, 13)
    NumPut("uchar", 0xF3, CLSID_PNG, 14)
    NumPut("uchar", 0x2E, CLSID_PNG, 15)
}

InitGDI() {
    global Canvas, hBitmap, hdcMem, G, CanvasWidth, CanvasHeight
    
    try {
        ControlGetPos(, , &w, &h, Canvas.Hwnd)
        CanvasWidth := w
        CanvasHeight := h
        
        hdc := DllCall("GetDC", "ptr", Canvas.Hwnd, "ptr")
        if !hdc {
            LogError("Failed to get device context")
            return false
        }
        
        hBitmap := DllCall("CreateCompatibleBitmap", "ptr", hdc, "int", CanvasWidth, "int", CanvasHeight, "ptr")
        if !hBitmap {
            DllCall("ReleaseDC", "ptr", Canvas.Hwnd, "ptr", hdc)
            LogError("Failed to create compatible bitmap")
            return false
        }
        
        hdcMem := DllCall("CreateCompatibleDC", "ptr", hdc, "ptr")
        if !hdcMem {
            DllCall("DeleteObject", "ptr", hBitmap)
            DllCall("ReleaseDC", "ptr", Canvas.Hwnd, "ptr", hdc)
            LogError("Failed to create compatible DC")
            return false
        }
        
        DllCall("SelectObject", "ptr", hdcMem, "ptr", hBitmap)
        
        result := DllCall("gdiplus\GdipCreateFromHDC", "ptr", hdcMem, "ptr*", &G, "int")
        if (result != 0) {
            DllCall("DeleteDC", "ptr", hdcMem)
            DllCall("DeleteObject", "ptr", hBitmap)
            DllCall("ReleaseDC", "ptr", Canvas.Hwnd, "ptr", hdc)
            LogError("GdipCreateFromHDC failed with error: " result)
            return false
        }
        
        DllCall("gdiplus\GdipSetSmoothingMode", "ptr", G, "int", 4)
        DllCall("ReleaseDC", "ptr", Canvas.Hwnd, "ptr", hdc)
        return true
        
    } catch as e {
        LogError("InitGDI exception: " e.Message)
        return false
    }
}

DrawGrid() {
    global G, CanvasWidth, CanvasHeight
    
    gridSize := Config.GRID_SIZE
    darkLineEvery := Config.GRID_MAJOR
    
    pGridPen := 0
    result := DllCall("gdiplus\GdipCreatePen1", "uint", 0xFFE0E0E0, "float", 1.0, "int", 2, "ptr*", &pGridPen, "int")
    if (result != 0)
        return
    
    x := gridSize
    while (x < CanvasWidth) {
        DllCall("gdiplus\GdipDrawLine", "ptr", G, "ptr", pGridPen, 
                "float", x, "float", 0, "float", x, "float", CanvasHeight)
        x += gridSize
    }
    
    y := gridSize
    while (y < CanvasHeight) {
        DllCall("gdiplus\GdipDrawLine", "ptr", G, "ptr", pGridPen, 
                "float", 0, "float", y, "float", CanvasWidth, "float", y)
        y += gridSize
    }
    
    DllCall("gdiplus\GdipDeletePen", "ptr", pGridPen)
    
    pDarkPen := 0
    result := DllCall("gdiplus\GdipCreatePen1", "uint", 0xFFC0C0C0, "float", 1.5, "int", 2, "ptr*", &pDarkPen, "int")
    if (result != 0)
        return
    
    x := darkLineEvery
    while (x < CanvasWidth) {
        DllCall("gdiplus\GdipDrawLine", "ptr", G, "ptr", pDarkPen, 
                "float", x, "float", 0, "float", x, "float", CanvasHeight)
        x += darkLineEvery
    }
    
    y := darkLineEvery
    while (y < CanvasHeight) {
        DllCall("gdiplus\GdipDrawLine", "ptr", G, "ptr", pDarkPen, 
                "float", 0, "float", y, "float", CanvasWidth, "float", y)
        y += darkLineEvery
    }
    
    DllCall("gdiplus\GdipDeletePen", "ptr", pDarkPen)
}

ScheduleRedraw() {
    global RedrawScheduled
    
    if !RedrawScheduled {
        RedrawScheduled := true
        SetTimer(DoRedraw, -16)
    }
}

DoRedraw() {
    global RedrawScheduled
    RedrawScheduled := false
    RedrawCanvas()
}

RedrawCanvas() {
    global G
    
    DllCall("gdiplus\GdipGraphicsClear", "ptr", G, "uint", 0xFFFFFFFF)
    DrawGrid()
    
    for index, connection in Connections
        DrawConnection(connection, index, G)
    
    for id, shape in Shapes
        DrawShape(shape, G)
    
    UpdateCanvasDisplay()
    UpdateStatus()
}

UpdateCanvasDisplay() {
    global Canvas, hdcMem, CanvasWidth, CanvasHeight
    
    hdc := DllCall("GetDC", "ptr", Canvas.Hwnd, "ptr")
    DllCall("BitBlt", "ptr", hdc, "int", 0, "int", 0, 
            "int", CanvasWidth, "int", CanvasHeight, 
            "ptr", hdcMem, "int", 0, "int", 0, "uint", 0x00CC0020)
    DllCall("ReleaseDC", "ptr", Canvas.Hwnd, "ptr", hdc)
}

GetCanvasCoords(&x, &y) {
    global Canvas, CanvasWidth, CanvasHeight
    
    MouseGetPos(&screenX, &screenY)
    ControlGetPos(&ctrlX, &ctrlY, , , Canvas.Hwnd)
    
    x := screenX - ctrlX
    y := screenY - ctrlY
    
    x := Max(0, Min(x, CanvasWidth - 1))
    y := Max(0, Min(y, CanvasHeight - 1))
}

LogError(message) {
    OutputDebug("ERROR: " message)
}

LogWarning(message) {
    OutputDebug("WARNING: " message)
}

ValidateShape(shape) {
    if !IsObject(shape)
        return false
    
    required := ["id", "type", "x", "y", "width", "height", "text", "color"]
    for field in required {
        if !shape.Has(field)
            return false
    }
    
    return true
}

ValidateInput(text) {
    ; Block potentially dangerous characters
    bannedChars := ["|", "`0", "`t`t", "``"]
    for char in bannedChars {
        if InStr(text, char)
            return false
    }
    
    ; Limit length
    if StrLen(text) > Config.MAX_TEXT_LENGTH
        return false
    
    return true
}

EscapeINI(text) {
    text := StrReplace(text, "\", "\\")
    text := StrReplace(text, "`n", "\n")
    text := StrReplace(text, "`r", "\r")
    text := StrReplace(text, "=", "\eq")
    text := StrReplace(text, "[", "\ob")
    text := StrReplace(text, "]", "\cb")
    text := StrReplace(text, ";", "\sc")
    text := StrReplace(text, "#", "\hs")
    return text
}

UnescapeINI(text) {
    text := StrReplace(text, "\n", "`n")
    text := StrReplace(text, "\r", "`r")
    text := StrReplace(text, "\eq", "=")
    text := StrReplace(text, "\ob", "[")
    text := StrReplace(text, "\cb", "]")
    text := StrReplace(text, "\sc", ";")
    text := StrReplace(text, "\hs", "#")
    text := StrReplace(text, "\\", "\")
    return text
}

SelectShapeType(shapeType) {
    global CurrentShapeType, IsDrawingConnection, CurrentSelectionText, SelectedConnection
    CurrentShapeType := shapeType
    IsDrawingConnection := false
    SelectedConnection := ""
    CurrentSelectionText.Text := shapeType
    UpdateStatus("Click on grid to place " shapeType)
}

ConnectTool(*) {
    global IsDrawingConnection, CurrentShapeType, StartShape, CurrentSelectionText, SelectedConnection
    IsDrawingConnection := true
    CurrentShapeType := ""
    StartShape := ""
    SelectedConnection := ""
    CurrentSelectionText.Text := "Connect Mode"
    UpdateStatus("Right-click first shape, then second shape to connect")
}

DisconnectTool(*) {
    global SelectedConnection, CurrentSelectionText
    if SelectedConnection {
        RemoveConnection(SelectedConnection)
        SelectedConnection := ""
        CurrentSelectionText.Text := "None"
        UpdateStatus("Connection removed")
    } else {
        CurrentSelectionText.Text := "Disconnect Mode"
        UpdateStatus("Double-click on a connection to remove it")
    }
}

CanvasClick(Ctrl, Info) {
    global CurrentShapeType, IsDrawingConnection, SelectedShape, IsDragging
    global DragOffsetX, DragOffsetY, Shapes, CurrentSelectionText, SelectedConnection
    
    GetCanvasCoords(&canvasX, &canvasY)
    
    connection := GetConnectionAtPos(canvasX, canvasY)
    if connection {
        SelectedConnection := connection
        SelectedShape := ""
        CurrentSelectionText.Text := "Connection selected"
        UpdateStatus("Connection selected - Click Disconnect button or double-click to remove")
        ScheduleRedraw()
        return
    }
    
    shape := GetShapeAtPos(canvasX, canvasY)
    
    if shape {
        SelectedShape := shape["id"]
        SelectedConnection := ""
        IsDragging := true
        DragOffsetX := canvasX - shape["x"]
        DragOffsetY := canvasY - shape["y"]
        CurrentSelectionText.Text := shape["text"]
        UpdateStatus("Dragging: " shape["text"])
        ScheduleRedraw()
    }
    else if CurrentShapeType {
        gridSize := Config.GRID_SIZE
        gridX := Round(canvasX / gridSize) * gridSize
        gridY := Round(canvasY / gridSize) * gridSize
        AddShape(CurrentShapeType, gridX - 60, gridY - 30)
    }
    else {
        SelectedShape := ""
        SelectedConnection := ""
        IsDragging := false
        CurrentSelectionText.Text := "None"
        UpdateStatus("Ready - Click shapes to select, drag to move")
        ScheduleRedraw()
    }
}

CanvasDoubleClick(Ctrl, Info) {
    GetCanvasCoords(&canvasX, &canvasY)
    
    connection := GetConnectionAtPos(canvasX, canvasY)
    if connection {
        RemoveConnection(connection)
        UpdateStatus("Connection removed")
    }
}

CheckMouse() {
    global IsDragging, SelectedShape, Shapes, CanvasWidth, CanvasHeight
    global DragOffsetX, DragOffsetY
    
    if !IsDragging || !SelectedShape || !Shapes.Has(SelectedShape)
        return
    
    if !WinActive("Flowdraw")
        return
    
    static lastX := 0, lastY := 0
    
    x := 0, y := 0
    GetCanvasCoords(&x, &y)
    
    if (x = lastX && y = lastY)
        return
    
    lastX := x
    lastY := y
    
    gridSize := Config.GRID_SIZE
    newX := Round((x - DragOffsetX) / gridSize) * gridSize
    newY := Round((y - DragOffsetY) / gridSize) * gridSize
    
    shape := Shapes[SelectedShape]
    newX := Max(0, Min(newX, CanvasWidth - shape["width"]))
    newY := Max(0, Min(newY, CanvasHeight - shape["height"]))
    
    ; Add undo step for first movement
    static lastSavedX := 0, lastSavedY := 0
    if (lastSavedX != shape["x"] || lastSavedY != shape["y"]) {
        lastSavedX := shape["x"]
        lastSavedY := shape["y"]
        AddUndoStep("move_shape", Map(
            "id", SelectedShape,
            "oldX", lastSavedX,
            "oldY", lastSavedY,
            "newX", newX,
            "newY", newY
        ))
    }
    
    Shapes[SelectedShape]["x"] := newX
    Shapes[SelectedShape]["y"] := newY
    
    ScheduleRedraw()
}

GuiContextMenu(GuiObj, GuiCtrl, *) {
    global IsDrawingConnection, StartShape, Connections, Shapes, StatsText
    
    if !IsDrawingConnection
        return
    
    GetCanvasCoords(&canvasX, &canvasY)
    
    shape := GetShapeAtPos(canvasX, canvasY)
    if !shape {
        UpdateStatus("No shape under cursor")
        return
    }
    
    if !StartShape {
        StartShape := shape["id"]
        UpdateStatus("First shape selected - Right-click second shape")
    } else if shape["id"] != StartShape {
        connection := Map("fromShape", StartShape, "toShape", shape["id"])
        Connections.Push(connection)
        
        AddUndoStep("add_connection", Map(
            "connection", connection.Clone()
        ))
        
        StatsText.Text := "Shapes: " Shapes.Count "`nConnections: " Connections.Length
        
        StartShape := ""
        IsDrawingConnection := false
        ScheduleRedraw()
        UpdateStatus("Shapes connected!")
    } else {
        UpdateStatus("Cannot connect shape to itself")
    }
}

AddShape(shapeType, x, y) {
    global ShapeCount, Shapes, SelectedShape, CurrentSelectionText, StatsText
    global CanvasWidth, CanvasHeight
    
    w := Config.SHAPE_DEFAULT_WIDTH
    h := Config.SHAPE_DEFAULT_HEIGHT
    
    x := Max(0, Min(x, CanvasWidth - w))
    y := Max(0, Min(y, CanvasHeight - h))
    
    ShapeCount++
    shapeId := "shape_" ShapeCount
    
    shape := Map(
        "id", shapeId,
        "type", shapeType,
        "x", x,
        "y", y,
        "width", w,
        "height", h,
        "text", GetDefaultText(shapeType),
        "color", GetDefaultColor(shapeType)
    )
    
    Shapes[shapeId] := shape
    SelectedShape := shapeId
    CurrentSelectionText.Text := shape["text"]
    
    AddUndoStep("add_shape", Map(
        "id", shapeId,
        "shape", shape.Clone()
    ))
    
    StatsText.Text := "Shapes: " Shapes.Count "`nConnections: " Connections.Length
    
    ScheduleRedraw()
    UpdateStatus("Added: " shape["text"] " - Drag to move")
}

GetDefaultText(shapeType) {
    switch shapeType {
        case "rectangle": return "Process"
        case "diamond": return "Decision?"
        case "ellipse": return "Start/End"
        case "data": return "Data"
        case "document": return "Document"
        default: return "Shape"
    }
}

GetDefaultColor(shapeType) {
    switch shapeType {
        case "rectangle": return "87CEEB"
        case "diamond": return "98FB98"
        case "ellipse": return "FFB6C1"
        case "data": return "FFFACD"
        case "document": return "DDA0DD"
        default: return "FFFFFF"
    }
}

GetShapeAtPos(x, y) {
    global Shapes
    
    ids := []
    for id in Shapes
        ids.Push(id)
    
    Loop ids.Length {
        id := ids[ids.Length - A_Index + 1]
        if !Shapes.Has(id)
            continue
        shape := Shapes[id]
        if PointInShape(x, y, shape)
            return shape
    }
    return ""
}

PointInShape(x, y, shape) {
    return (x >= shape["x"] && x <= shape["x"] + shape["width"]
        && y >= shape["y"] && y <= shape["y"] + shape["height"])
}

EditText(*) {
    global SelectedShape, Shapes, CurrentSelectionText
    
    if !SelectedShape || !Shapes.Has(SelectedShape) {
        MsgBox("Please select a shape first!", "No Selection", "Icon! 48")
        return
    }
    
    shape := Shapes[SelectedShape]
    maxLen := Config.MAX_TEXT_LENGTH
    
    result := InputBox(
        "Enter text for " shape["type"] ":`n(Maximum " maxLen " characters)", 
        "Edit Text", 
        "w400 h150", 
        shape["text"]
    )
    
    if result.Result = "OK" && result.Value != "" {
        if !ValidateInput(result.Value) {
            MsgBox("Invalid input!`n`nPlease avoid using special characters like | or ``", "Input Error", "Icon! 16")
            return
        }
        
        newText := SubStr(result.Value, 1, maxLen)
        
        if StrLen(result.Value) > maxLen {
            MsgBox("Text was truncated to " maxLen " characters.", "Warning", "Icon! 48")
        }
        
        AddUndoStep("edit_text", Map(
            "id", SelectedShape,
            "oldText", shape["text"],
            "newText", newText
        ))
        
        shape["text"] := newText
        CurrentSelectionText.Text := newText
        ScheduleRedraw()
        UpdateStatus("Text updated: " newText)
    }
}

DeleteSelected(*) {
    global SelectedShape, Shapes, CurrentSelectionText, StatsText, SelectedConnection
    
    if SelectedShape && Shapes.Has(SelectedShape) {
        shape := Shapes[SelectedShape]
        shapeName := shape["text"]
        
        AddUndoStep("delete_shape", Map(
            "id", SelectedShape,
            "shape", shape.Clone()
        ))
        
        Shapes.Delete(SelectedShape)
        RemoveConnections(SelectedShape)
        SelectedShape := ""
        CurrentSelectionText.Text := "None"
        
        StatsText.Text := "Shapes: " Shapes.Count "`nConnections: " Connections.Length
        
        ScheduleRedraw()
        UpdateStatus("Shape deleted: " shapeName)
    } 
    else if SelectedConnection {
        connection := Connections[SelectedConnection]
        AddUndoStep("delete_connection", Map(
            "connection", connection.Clone(),
            "index", SelectedConnection
        ))
        RemoveConnection(SelectedConnection)
        UpdateStatus("Connection deleted")
    }
    else {
        MsgBox("Please select a shape or connection first!", "No Selection", "Icon! 48")
    }
}

RemoveConnections(shapeId) {
    global Connections, StatsText
    
    i := 1
    removed := 0
    while i <= Connections.Length {
        conn := Connections[i]
        if (conn["fromShape"] = shapeId || conn["toShape"] = shapeId) {
            Connections.RemoveAt(i)
            removed++
        } else {
            i++
        }
    }
    
    StatsText.Text := "Shapes: " Shapes.Count "`nConnections: " Connections.Length
    
    if removed > 0
        LogWarning("Removed " removed " connections for shape: " shapeId)
}

ClearAll(*) {
    global Shapes, Connections, SelectedShape, ShapeCount, CurrentSelectionText, StatsText, SelectedConnection
    
    if MsgBox("Clear everything?`n`nThis action cannot be undone.", "Confirm", "YesNo Icon! 32") = "Yes" {
        AddUndoStep("clear_all", Map(
            "shapes", Shapes.Clone(),
            "connections", Connections.Clone()
        ))
        
        Shapes := Map()
        Connections := Array()
        SelectedShape := ""
        SelectedConnection := ""
        ShapeCount := 0
        CurrentSelectionText.Text := "None"
        StatsText.Text := "Shapes: 0`nConnections: 0"
        ScheduleRedraw()
        UpdateStatus("Canvas cleared")
    }
}

GetConnectionAtPos(x, y) {
    global Connections, Shapes
    
    for index, connection in Connections {
        if !Shapes.Has(connection["fromShape"]) || !Shapes.Has(connection["toShape"])
            continue
        
        from := Shapes[connection["fromShape"]]
        to := Shapes[connection["toShape"]]
        
        startX := from["x"] + from["width"]
        startY := from["y"] + from["height"] / 2
        endX := to["x"]
        endY := to["y"] + to["height"] / 2
        
        if PointNearLine(x, y, startX, startY, endX, endY, 5) {
            return index
        }
    }
    return ""
}

PointNearLine(px, py, x1, y1, x2, y2, tolerance) {
    A := px - x1
    B := py - y1
    C := x2 - x1
    D := y2 - y1
    
    dot := A * C + B * D
    len_sq := C * C + D * D
    
    if (len_sq = 0)
        return false
    
    param := dot / len_sq
    
    if (param < 0) {
        xx := x1
        yy := y1
    }
    else if (param > 1) {
        xx := x2
        yy := y2
    }
    else {
        xx := x1 + param * C
        yy := y1 + param * D
    }
    
    dx := px - xx
    dy := py - yy
    distance := Sqrt(dx * dx + dy * dy)
    
    return distance <= tolerance
}

RemoveConnection(connectionIndex) {
    global Connections, StatsText, SelectedConnection
    
    if !connectionIndex || connectionIndex > Connections.Length {
        LogWarning("Invalid connection index: " connectionIndex)
        return false
    }
    
    try {
        Connections.RemoveAt(connectionIndex)
        SelectedConnection := ""
        
        StatsText.Text := "Shapes: " Shapes.Count "`nConnections: " Connections.Length
        
        ScheduleRedraw()
        return true
    } catch as e {
        LogError("RemoveConnection failed: " e.Message)
        return false
    }
}

DrawShape(shape, pGraphics) {
    global SelectedShape
    
    if !ValidateShape(shape) {
        LogWarning("Invalid shape object")
        return
    }
    
    fillColor := "0xFF" shape["color"]
    borderColor := 0xFF000000
    textColor := 0xFF000000
    
    pBrush := 0
    pPen := 0
    pSelectionPen := 0
    
    try {
        result := DllCall("gdiplus\GdipCreateSolidFill", "uint", fillColor, "ptr*", &pBrush, "int")
        if (result != 0)
            throw Error("Failed to create brush")
        
        result := DllCall("gdiplus\GdipCreatePen1", "uint", borderColor, "float", 2.0, "int", 2, "ptr*", &pPen, "int")
        if (result != 0)
            throw Error("Failed to create pen")
        
        if shape["id"] = SelectedShape {
            result := DllCall("gdiplus\GdipCreatePen1", "uint", 0xFFFF0000, "float", 3.0, "int", 2, "ptr*", &pSelectionPen, "int")
            if (result = 0) {
                DllCall("gdiplus\GdipDrawRectangle", "ptr", pGraphics, "ptr", pSelectionPen, 
                        "float", shape["x"]-2, "float", shape["y"]-2, 
                        "float", shape["width"]+4, "float", shape["height"]+4)
            }
        }
        
        switch shape["type"] {
            case "rectangle":
                DllCall("gdiplus\GdipFillRectangle", "ptr", pGraphics, "ptr", pBrush, 
                        "float", shape["x"], "float", shape["y"], 
                        "float", shape["width"], "float", shape["height"])
                DllCall("gdiplus\GdipDrawRectangle", "ptr", pGraphics, "ptr", pPen, 
                        "float", shape["x"], "float", shape["y"], 
                        "float", shape["width"], "float", shape["height"])
            
            case "ellipse":
                DllCall("gdiplus\GdipFillEllipse", "ptr", pGraphics, "ptr", pBrush, 
                        "float", shape["x"], "float", shape["y"], 
                        "float", shape["width"], "float", shape["height"])
                DllCall("gdiplus\GdipDrawEllipse", "ptr", pGraphics, "ptr", pPen, 
                        "float", shape["x"], "float", shape["y"], 
                        "float", shape["width"], "float", shape["height"])
            
            case "diamond":
                DrawPolygonShape(shape, pPen, pBrush, pGraphics, "diamond")
            
            case "data":
                DrawPolygonShape(shape, pPen, pBrush, pGraphics, "parallelogram")
            
            case "document":
                DrawPolygonShape(shape, pPen, pBrush, pGraphics, "document")
        }
        
        DrawTextOnShape(shape, textColor, pGraphics)
        
    } catch as e {
        LogError("DrawShape error: " e.Message)
    } finally {
        ; Proper resource cleanup - CRITICAL FIX
        if pBrush
            DllCall("gdiplus\GdipDeleteBrush", "ptr", pBrush)
        if pPen
            DllCall("gdiplus\GdipDeletePen", "ptr", pPen)
        if pSelectionPen
            DllCall("gdiplus\GdipDeletePen", "ptr", pSelectionPen)
    }
}

DrawPolygonShape(shape, pPen, pBrush, pGraphics, shapeType) {
    points := ""
    pointCount := 0
    
    switch shapeType {
        case "diamond":
            points := Buffer(32)
            pointCount := 4
            cx := shape["x"] + shape["width"]/2
            cy := shape["y"] + shape["height"]/2
            NumPut("float", cx, points, 0)
            NumPut("float", shape["y"], points, 4)
            NumPut("float", shape["x"] + shape["width"], points, 8)
            NumPut("float", cy, points, 12)
            NumPut("float", cx, points, 16)
            NumPut("float", shape["y"] + shape["height"], points, 20)
            NumPut("float", shape["x"], points, 24)
            NumPut("float", cy, points, 28)
        
        case "parallelogram":
            points := Buffer(32)
            pointCount := 4
            offset := 20
            NumPut("float", shape["x"] + offset, points, 0)
            NumPut("float", shape["y"], points, 4)
            NumPut("float", shape["x"] + shape["width"], points, 8)
            NumPut("float", shape["y"], points, 12)
            NumPut("float", shape["x"] + shape["width"] - offset, points, 16)
            NumPut("float", shape["y"] + shape["height"], points, 20)
            NumPut("float", shape["x"], points, 24)
            NumPut("float", shape["y"] + shape["height"], points, 28)
        
        case "document":
            points := Buffer(40)
            pointCount := 5
            w := shape["width"]
            h := shape["height"]
            x := shape["x"]
            y := shape["y"]
            fold := 15
            NumPut("float", x, points, 0)
            NumPut("float", y, points, 4)
            NumPut("float", x + w, points, 8)
            NumPut("float", y, points, 12)
            NumPut("float", x + w, points, 16)
            NumPut("float", y + h - fold, points, 20)
            NumPut("float", x + w - fold, points, 24)
            NumPut("float", y + h, points, 28)
            NumPut("float", x, points, 32)
            NumPut("float", y + h, points, 36)
    }
    
    if points && pointCount > 0 {
        DllCall("gdiplus\GdipFillPolygon", "ptr", pGraphics, "ptr", pBrush, "ptr", points, "int", pointCount, "int", 0)
        DllCall("gdiplus\GdipDrawPolygon", "ptr", pGraphics, "ptr", pPen, "ptr", points, "int", pointCount)
    }
}

DrawTextOnShape(shape, textColor, pGraphics) {
    hFamily := 0
    hFont := 0
    hFormat := 0
    hBrush := 0
    
    try {
        result := DllCall("gdiplus\GdipCreateFontFamilyFromName", "wstr", "Arial", "ptr", 0, "ptr*", &hFamily, "int")
        if (result != 0)
            return
        
        DllCall("gdiplus\GdipCreateFont", "ptr", hFamily, "float", 11, "int", 0, "int", 0, "ptr*", &hFont)
        DllCall("gdiplus\GdipCreateStringFormat", "int", 0, "int", 0, "ptr*", &hFormat)
        DllCall("gdiplus\GdipSetStringFormatAlign", "ptr", hFormat, "int", 1)
        DllCall("gdiplus\GdipSetStringFormatLineAlign", "ptr", hFormat, "int", 1)
        DllCall("gdiplus\GdipCreateSolidFill", "uint", textColor, "ptr*", &hBrush)
        
        rect := Buffer(16)
        NumPut("float", shape["x"], rect, 0)
        NumPut("float", shape["y"], rect, 4)
        NumPut("float", shape["width"], rect, 8)
        NumPut("float", shape["height"], rect, 12)
        
        DllCall("gdiplus\GdipDrawString", "ptr", pGraphics, "wstr", shape["text"], 
                "int", -1, "ptr", hFont, "ptr", rect, "ptr", hFormat, "ptr", hBrush)
        
    } catch as e {
        LogError("DrawTextOnShape error: " e.Message)
    } finally {
        ; Proper resource cleanup - CRITICAL FIX
        if hBrush
            DllCall("gdiplus\GdipDeleteBrush", "ptr", hBrush)
        if hFormat
            DllCall("gdiplus\GdipDeleteStringFormat", "ptr", hFormat)
        if hFont
            DllCall("gdiplus\GdipDeleteFont", "ptr", hFont)
        if hFamily
            DllCall("gdiplus\GdipDeleteFontFamily", "ptr", hFamily)
    }
}

DrawConnection(connection, index, pGraphics) {
    global Shapes, SelectedConnection
    
    if !Shapes.Has(connection["fromShape"]) || !Shapes.Has(connection["toShape"])
        return
    
    from := Shapes[connection["fromShape"]]
    to := Shapes[connection["toShape"]]
    
    startX := from["x"] + from["width"]
    startY := from["y"] + from["height"] / 2
    endX := to["x"]
    endY := to["y"] + to["height"] / 2
    
    penColor := (index = SelectedConnection) ? 0xFFFF0000 : 0xFF000000
    penWidth := (index = SelectedConnection) ? 3.0 : 2.0
    
    pPen := 0
    result := DllCall("gdiplus\GdipCreatePen1", "uint", penColor, "float", penWidth, "int", 2, "ptr*", &pPen, "int")
    
    if (result = 0) {
        DllCall("gdiplus\GdipDrawLine", "ptr", pGraphics, "ptr", pPen, 
                "float", startX, "float", startY, "float", endX, "float", endY)
        DllCall("gdiplus\GdipDeletePen", "ptr", pPen) ; CRITICAL FIX - Cleanup pen
    }
}

SaveProjectINI(*) {
    global Shapes, Connections
    
    selectedFile := FileSelect("S16", "MyFlowdraw.ini", "Save Project", "INI Files (*.ini)")
    if !selectedFile
        return
    
    if !RegExMatch(selectedFile, "\.ini$")
        selectedFile .= ".ini"
    
    if InStr(selectedFile, "..") {
        MsgBox("Invalid file path detected!`n`nPath traversal is not allowed.", "Security Error", "Icon! 16")
        return
    }
    
    try {
        iniContent := ""
        
        shapeIndex := 0
        for id, shape in Shapes {
            if !ValidateShape(shape) {
                LogWarning("Skipping invalid shape: " id)
                continue
            }
            
            shapeIndex++
            section := "Shape" shapeIndex
            iniContent .= "[" section "]`n"
            iniContent .= "ID=" id "`n"
            iniContent .= "Type=" shape["type"] "`n"
            iniContent .= "X=" shape["x"] "`n"
            iniContent .= "Y=" shape["y"] "`n"
            iniContent .= "Width=" shape["width"] "`n"
            iniContent .= "Height=" shape["height"] "`n"
            iniContent .= "Text=" EscapeINI(shape["text"]) "`n"
            iniContent .= "Color=" shape["color"] "`n`n"
        }
        
        connectionIndex := 0
        for index, connection in Connections {
            if !Shapes.Has(connection["fromShape"]) || !Shapes.Has(connection["toShape"]) {
                LogWarning("Skipping invalid connection: " index)
                continue
            }
            
            connectionIndex++
            section := "Connection" connectionIndex
            iniContent .= "[" section "]`n"
            iniContent .= "From=" connection["fromShape"] "`n"
            iniContent .= "To=" connection["toShape"] "`n`n"
        }
        
        iniContent .= "[Info]`n"
        iniContent .= "Version=2.0`n"
        iniContent .= "TotalShapes=" shapeIndex "`n"
        iniContent .= "TotalConnections=" connectionIndex "`n"
        iniContent .= "Saved=" A_Now "`n"
        iniContent .= "CanvasWidth=" CanvasWidth "`n"
        iniContent .= "CanvasHeight=" CanvasHeight "`n"
        
        file := FileOpen(selectedFile, "w", "UTF-8-RAW")
        if !file {
            throw Error("Failed to open file for writing")
        }
        
        bytesWritten := file.Write(iniContent)
        file.Close()
        
        if bytesWritten = 0 {
            throw Error("Failed to write data to file")
        }
        
        MsgBox("Project saved successfully!`n`nFile: " selectedFile "`nShapes: " shapeIndex "`nConnections: " connectionIndex, 
               "Save Successful", "Iconi 64")
        UpdateStatus("Project saved: " selectedFile)
        
    } catch as e {
        MsgBox("Save failed!`n`nError: " e.Message "`n`nPlease check if you have write permissions.", 
               "Save Error", "Icon! 16")
        LogError("SaveProjectINI failed: " e.Message)
    }
}

LoadProjectINI(*) {
    global Shapes, Connections, SelectedShape, SelectedConnection, StatsText, CurrentSelectionText, ShapeCount
    
    selectedFile := FileSelect("3", , "Load Project", "INI Files (*.ini)")
    if !selectedFile
        return
    
    if !FileExist(selectedFile) {
        MsgBox("File does not exist!`n`n" selectedFile, "Load Error", "Icon! 16")
        return
    }
    
    try {
        Shapes := Map()
        Connections := Array()
        SelectedShape := ""
        SelectedConnection := ""
        ShapeCount := 0
        
        version := IniRead(selectedFile, "Info", "Version", "1.0")
        totalShapes := Integer(IniRead(selectedFile, "Info", "TotalShapes", "0"))
        totalConnections := Integer(IniRead(selectedFile, "Info", "TotalConnections", "0"))
        
        if totalShapes = 0 {
            MsgBox("No shapes found in file!`n`nThe file might be empty or corrupted.", 
                   "Load Warning", "Icon! 48")
        }
        
        shapesLoaded := 0
        Loop totalShapes {
            section := "Shape" A_Index
            
            id := IniRead(selectedFile, section, "ID", "")
            if !id {
                LogWarning("Shape " A_Index " has no ID, skipping")
                continue
            }
            
            try {
                shape := Map()
                shape["id"] := id
                shape["type"] := IniRead(selectedFile, section, "Type", "rectangle")
                
                xVal := IniRead(selectedFile, section, "X", "0")
                yVal := IniRead(selectedFile, section, "Y", "0")
                wVal := IniRead(selectedFile, section, "Width", "120")
                hVal := IniRead(selectedFile, section, "Height", "60")
                
                if !IsNumber(xVal) || !IsNumber(yVal) || !IsNumber(wVal) || !IsNumber(hVal) {
                    throw ValueError("Invalid numeric values")
                }
                
                shape["x"] := Integer(xVal)
                shape["y"] := Integer(yVal)
                shape["width"] := Integer(wVal)
                shape["height"] := Integer(hVal)
                
                textVal := IniRead(selectedFile, section, "Text", "Shape")
                shape["text"] := UnescapeINI(textVal)
                
                shape["color"] := IniRead(selectedFile, section, "Color", "FFFFFF")
                
                if ValidateShape(shape) {
                    Shapes[id] := shape
                    shapesLoaded++
                    
                    if RegExMatch(id, "shape_(\d+)", &match)
                        ShapeCount := Max(ShapeCount, Integer(match[1]))
                } else {
                    LogWarning("Shape " id " failed validation")
                }
                
            } catch as e {
                LogError("Failed to load shape " A_Index ": " e.Message)
                continue
            }
        }
        
        connectionsLoaded := 0
        Loop totalConnections {
            section := "Connection" A_Index
            
            fromShape := IniRead(selectedFile, section, "From", "")
            toShape := IniRead(selectedFile, section, "To", "")
            
            if !fromShape || !toShape {
                LogWarning("Connection " A_Index " has missing data")
                continue
            }
            
            if !Shapes.Has(fromShape) || !Shapes.Has(toShape) {
                LogWarning("Connection " A_Index " references non-existent shapes")
                continue
            }
            
            connection := Map("fromShape", fromShape, "toShape", toShape)
            Connections.Push(connection)
            connectionsLoaded++
        }
        
        CurrentSelectionText.Text := "None"
        StatsText.Text := "Shapes: " Shapes.Count "`nConnections: " Connections.Length
        
        ScheduleRedraw()
        
        msg := "Project loaded successfully!`n`n"
        msg .= "File: " selectedFile "`n"
        msg .= "Version: " version "`n"
        msg .= "Shapes loaded: " shapesLoaded " of " totalShapes "`n"
        msg .= "Connections loaded: " connectionsLoaded " of " totalConnections
        
        if shapesLoaded < totalShapes || connectionsLoaded < totalConnections {
            msg .= "`n`n⚠️ Warning: Some items could not be loaded.`nCheck the debug output for details."
        }
        
        MsgBox(msg, "Load Successful", "Iconi 64")
        UpdateStatus("Project loaded: " shapesLoaded " shapes, " connectionsLoaded " connections")
        
    } catch as e {
        MsgBox("Load failed!`n`nError: " e.Message "`n`nThe file might be corrupted or in an unsupported format.", 
               "Load Error", "Icon! 16")
        LogError("LoadProjectINI failed: " e.Message)
    }
}

ExportToPNG(*) {
    global Shapes, Connections, CanvasWidth, CanvasHeight, CLSID_PNG
    
    selectedFile := FileSelect("S16", "MyFlowdraw.png", "Export to PNG", "PNG Files (*.png)")
    if !selectedFile
        return
    
    if !RegExMatch(selectedFile, "\.png$")
        selectedFile .= ".png"
    
    if InStr(selectedFile, "..") {
        MsgBox("Invalid file path detected!`n`nPath traversal is not allowed.", "Security Error", "Icon! 16")
        return
    }
    
    pBitmap := 0
    pGraphics := 0
    
    try {
        result := DllCall("gdiplus\GdipCreateBitmapFromScan0", 
                         "int", CanvasWidth, "int", CanvasHeight, 
                         "int", 0, "int", 0x26200A, "ptr", 0, "ptr*", &pBitmap, "int")
        
        if (result != 0) {
            throw Error("Failed to create bitmap (Error: " result ")")
        }
        
        result := DllCall("gdiplus\GdipGetImageGraphicsContext", "ptr", pBitmap, "ptr*", &pGraphics, "int")
        
        if (result != 0) {
            throw Error("Failed to create graphics context (Error: " result ")")
        }
        
        DllCall("gdiplus\GdipSetSmoothingMode", "ptr", pGraphics, "int", 4)
        DllCall("gdiplus\GdipGraphicsClear", "ptr", pGraphics, "uint", 0xFFFFFFFF)
        
        for index, connection in Connections
            DrawConnection(connection, 0, pGraphics)
        
        for id, shape in Shapes
            DrawShape(shape, pGraphics)
        
        result := DllCall("gdiplus\GdipSaveImageToFile", 
                         "ptr", pBitmap, "wstr", selectedFile, 
                         "ptr", CLSID_PNG.Ptr, "ptr", 0, "int")
        
        if (result != 0) {
            throw Error("GDI+ save failed (Error: " result ")")
        }
        
        if !FileExist(selectedFile) {
            throw Error("File was not created")
        }
        
        fileSize := FileGetSize(selectedFile)
        if fileSize < 100 {
            throw Error("File is too small (" fileSize " bytes), export likely failed")
        }
        
        MsgBox("Flowdraw exported successfully!`n`nFile: " selectedFile "`nSize: " FormatBytes(fileSize), 
               "Export Successful", "Iconi 64")
        UpdateStatus("Exported to: " selectedFile)
        
    } catch as e {
        MsgBox("Export failed!`n`nError: " e.Message "`n`nPlease check write permissions and disk space.", 
               "Export Error", "Icon! 16")
        LogError("ExportToPNG failed: " e.Message)
    } finally {
        if pGraphics
            DllCall("gdiplus\GdipDeleteGraphics", "ptr", pGraphics)
        if pBitmap
            DllCall("gdiplus\GdipDisposeImage", "ptr", pBitmap)
    }
}

FormatBytes(bytes) {
    if bytes < 1024
        return bytes " bytes"
    if bytes < 1048576
        return Round(bytes / 1024, 1) " KB"
    return Round(bytes / 1048576, 1) " MB"
}

UpdateStatus(extra := "") {
    global StatusText, Shapes, Connections, SelectedShape, SelectedConnection
    
    status := "Shapes: " Shapes.Count " | Connections: " Connections.Length 
    status .= " | Selected: " (SelectedShape && Shapes.Has(SelectedShape) ? Shapes[SelectedShape]["text"] : SelectedConnection ? "Connection" : "None")
    
    if extra != ""
        status .= " | " extra
    
    StatusText.Text := status
}

GuiClose(*) {
    global gpToken, hdcMem, hBitmap, G
    
    try {
        if G {
            DllCall("gdiplus\GdipDeleteGraphics", "ptr", G)
            G := 0
        }
        if hdcMem {
            DllCall("DeleteDC", "ptr", hdcMem)
            hdcMem := 0
        }
        if hBitmap {
            DllCall("DeleteObject", "ptr", hBitmap)
            hBitmap := 0
        }
        if gpToken {
            DllCall("gdiplus\GdiplusShutdown", "ptr", gpToken)
            gpToken := 0
        }
    } catch as e {
        LogError("Cleanup error: " e.Message)
    }
    
    ExitApp
}

; === HOTKEYS ===

#HotIf WinActive("Flowdraw")

^1::SelectShapeType("rectangle")
^2::SelectShapeType("ellipse") 
^3::SelectShapeType("diamond")
^c::ConnectTool()
^d::DisconnectTool()
Delete::DeleteSelected()
^l::ClearAll()
^s::SaveProjectINI()
^o::LoadProjectINI()
^e::ExportToPNG()
^z::UndoAction()
^y::RedoAction()

Esc::HandleEscape()

#HotIf

~LButton Up::HandleMouseUp()

; Helper functions for hotkeys
HandleEscape() {
    global IsDrawingConnection, StartShape, CurrentSelectionText, SelectedConnection
    if (IsDrawingConnection) {
        IsDrawingConnection := false
        StartShape := ""
        CurrentSelectionText.Text := "None"
        UpdateStatus("Connection mode cancelled")
    }
    else if (SelectedConnection) {
        SelectedConnection := ""
        CurrentSelectionText.Text := "None"
        ScheduleRedraw()
        UpdateStatus("Connection deselected")
    }
}

HandleMouseUp() {
    global IsDragging
    IsDragging := false
}