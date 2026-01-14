#Requires AutoHotkey v2

; Global variables
global isRunning := false
global stopRequested := false
global waitingForEnter := false
global enterPressed := false

; Esc key to force stop
~Esc:: {
    global stopRequested, isRunning
    if isRunning {
        stopRequested := true
        MsgBox "処理を中断します...", "中断", "T2"
    }
}

; Enter key detection
~Enter:: {
    global waitingForEnter, enterPressed
    if waitingForEnter {
        enterPressed := true
    }
}

; Main process
if isRunning {
    MsgBox "すでに実行中です", "警告"
    ExitApp
}

csvPath := FileSelect(3, , "CSVファイルを選択してください", "CSV Files (*.csv)")
if (csvPath = "") {
    MsgBox "キャンセルされました", "中止"
    ExitApp
}

if !FileExist(csvPath) {
    MsgBox "ファイルが見つかりません:`n" csvPath, "エラー"
    ExitApp
}

result := MsgBox("CSVファイル: " csvPath "`n`n最初のセルをクリックしてください。`n`n※処理中はEscキーで中断できます", "確認", "OKCancel")
if (result = "Cancel")
    ExitApp

ClickWait()
Sleep(300)

CloseErrorPopup()
Sleep(200)

isRunning := true
stopRequested := false

try {
    csvFile := FileRead(csvPath, "UTF-8")
    rows := StrSplit(csvFile, "`n")
    
    totalRows := 0
    for row in rows {
        if (Trim(row) != "")
            totalRows++
    }
    
    processedRows := 0
    
    for row in rows {
        if stopRequested {
            MsgBox "処理を中断しました。`n処理済み: " processedRows " / " totalRows " 行", "中断完了"
            break
        }
        
        row := Trim(row)
        if (row = "")
            continue
        
        cols := StrSplit(row, ",")
        
        if stopRequested
            break
        
        ; W/C input
        first := cols[1]
        if !RegExMatch(first, "^W", &out)
            first := "W" . first
        first := StrUpper(first)
        
        Send(first)
        Sleep(100)
        
        Sleep(200)
        if CheckForError() {
            MsgBox "エラーを検出しました。`n行番号: " (processedRows + 1) "`nデータ: " row, "エラー検出", "T5"
            break
        }
        
        Send("{Tab}")
        Sleep(100)
        
        ; Work number input
        if cols.Length >= 2 {
            Send(StrUpper(cols[2]))
            Sleep(100)
            
            Send("{Tab}")
            Sleep(500)
            
            ; Debug: Check what window is active
            activeTitle := WinGetTitle("A")
            
            if CheckForError() {
                ; Popup detected - let user handle it manually
                MsgBox("工番『" cols[2] "』が見つかりません。`n`n以下の手順で修正してください：`n1. ポップアップのOKをクリック`n2. 正しい工番を入力`n3. Enterキーを押す`n`nその後、処理を続行します。", "工番エラー - 手動修正", "OK")
                
                ToolTip "ポップアップOK → 工番修正 → Enter押下", 0, 0
                
                ; Wait for user to press Enter
                waitingForEnter := true
                enterPressed := false
                
                Loop {
                    if stopRequested {
                        ToolTip
                        waitingForEnter := false
                        break
                    }
                    if enterPressed {
                        break
                    }
                    Sleep(50)
                }
                
                ToolTip
                waitingForEnter := false
                
                if stopRequested {
                    break
                }
                
                Sleep(500)
                
                ; Input work time
                if cols.Length >= 3 {
                    workTime := Trim(cols[3])
                    Send(workTime)
                    Sleep(100)
                }
                
                ; Move to next row
                Send("{Tab}{Tab}{Tab}")
                Sleep(150)
                
                Send("+{Tab}+{Tab}+{Tab}")
                Sleep(150)
                
                processedRows++
                continue
            }
        }
        
        ; Work time input
        if cols.Length >= 3 {
            workTime := Trim(cols[3])
            Send(workTime)
            Sleep(100)
            
            if CheckForError() {
                MsgBox "エラーを検出しました（作業時間入力時）。`n行番号: " (processedRows + 1), "エラー検出", "T5"
                break
            }
        }
        
        ; Move to next row
        Send("{Tab}{Tab}{Tab}")
        Sleep(150)
        
        Send("+{Tab}+{Tab}+{Tab}")
        Sleep(150)
        
        processedRows++
    }
    
    if !stopRequested {
        MsgBox "すべての行の処理が完了しました。`n処理済み: " processedRows " / " totalRows " 行", "完了"
    }
    
} catch as err {
    MsgBox "エラーが発生しました:`n" err.Message, "エラー"
} finally {
    isRunning := false
    stopRequested := false
    waitingForEnter := false
}

ClickWait() {
    global stopRequested
    state := 0
    Loop {
        if stopRequested
            return
        if GetKeyState("LButton", "P") {
            state := 1
        }
        else if (state = 1) {
            return
        }
        Sleep(10)
    }
}

CheckForError() {
    ; Check for warning popup (title is "警告")
    if WinExist("警告") {
        return true
    }
    
    activeTitle := WinGetTitle("A")
    if (activeTitle = "警告") {
        return true
    }
    
    return false
}

CloseErrorPopup() {
    Loop 10 {
        if WinExist("警告") {
            WinActivate("警告")
            Sleep(100)
            Send("{Enter}")
            Sleep(200)
        } else {
            return true
        }
    }
    return false
}