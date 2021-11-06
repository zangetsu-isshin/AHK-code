#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#Persistent
#singleinstance, force
SetTitleMatchMode, 2

FormatTime, today, YYYYMMDDHH24MISS, M/dd/yyyy

port_no := IsChromeInDebugMode()

if (port_no = 9222)
{
    WinGet, chrome_count, List , Google Chrome
    if (chrome_count > 1)
        chrome_ahk_id := chrome_count1
    else
        WinGet, chrome_ahk_id, ID , Google Chrome
}
else
{
    msgbox, 16,Chrome Error, 
    (
        Chrome is not in Debug mode.
Google Chrome will now close and
reopen in Debug Mode.
    )
    WinKill, Google Chrome
    sleep 500
    run chrome.exe "--remote-debugging-port=9222"
    sleep 2000
    WinWait, Google Chrome
    sleep 1000
    reload
}
; SeleniumBasic 2.0.9.0

^!r::reload

#j::
    driver := ChromeGet()

    msgbox % driver.title

    Try driver.SwitchToNextWindow()
    Catch {

    }

    msgbox % driver.title

    driver := SwitchToChromeTab("tab 2", driver)

    msgbox % "returned driver.title : " driver.title

return

SwitchToChromeTab(desired_tab, driver, mode := "normal") {

    tab_found := False
    open_tabs := driver.Windows

    if (InStr(driver.title,desired_tab))
    {
        msgbox, Already on the tab
        return driver
    }

    for x in driver.Windows
    {
        msgbox, % "Tab " A_Index "`nTitle : " x.title "`n`ndriver.title : " driver.title

        if (instr(x.title, desired_tab))
        {
            msgbox, % "Found!`nx.title: " x.title
            tab_found := True

            break
        }

    }

    if not (tab_found)
    {
        msgbox, % desired_tab " was not found!"
        sleep 500
        Chrome_OpenNewTab()
        sleep 500
        x := ChromeGet()
        x.SwitchToNextWindow()
        sleep 500
        return x
    }

return driver

}

IsChromeInDebugMode() {
    WinGet, pid, PID, ahk_exe chrome.exe
    ; if !pid
    ;      throw "Chrome window not found"

    for item in ComObjGet("winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId='" pid "'")
        if RegExMatch(item.CommandLine, "i)--remote-debugging-port=\K\d+", port)
        return port
}

ChromeGet(IP_Port := "127.0.0.1:9222") {

    loop, {
        try, 
        {
            ; ToolTip, Starting ChromeDriver
            driver := ComObjCreate("Selenium.CHROMEDriver") ;Chrome driver

            driver.SetCapability("debuggerAddress", IP_Port)
            ; driver.SetCapability("test-type", "disable-popup-blocking")
            driver.Start()
            ; ToolTip
            break
        }
        catch e
        {
            ToolTip, Updating ChromeDriver
            driver := ""
            msgbox,,Driver start error, Updating chromedriver now. Kill all other running scripts and press OK.
            SChrome_UpdateDriver()
            ToolTip
        }

    }

return driver
}

Chrome_OpenNewTab(url := "", port := "9222") {

    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", "http://127.0.0.1:" port "/json/new?" url)
    whr.Send()
}

;; Helper functions
;; --

SChrome_UpdateDriver(mode := "na") {
    ; Function Updates the Chromedriver by checking the versions and downloading the latest chromedriver.
    ; Written by AHK_User
    ; Thanks to tmplinshi

    Dir_Chromedriver:= "C:\Users\" A_UserName "\AppData\Local\SeleniumBasic\chromedriver.exe"

    SplitPath, Dir_Chromedriver, , Folder_Chromedriver

    WinGet, chrome_path, ProcessPath, % "ahk_exe chrome.exe"	; activePath is the output variable and can be named anything you like, ProcessPath is a fixed parameter, specifying the action of the winget command.
    FileGetVersion, Version_Chrome, %chrome_path% ; In case chrome is not in program file
    ; FileGetVersion, Version_Chrome, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe    ; In case chrome is in program file

    Version_Chrome := RegexReplace(Version_Chrome, "\.\d+$")

    ; msgbox, Folder = %Folder_Chromedriver%
    ; if (mode != "auto_check")
    ; msgbox, Chrome Version : %Version_Chrome%
    ; Get Chromedriver version
    Version_ChromeDriver := RunHide("""" Dir_Chromedriver """ --version")
    ;~ DebugWindow("`nVersion Chromedriver:" Version_Chromedriver,Clear:=0,LineBreak:=1)
    Version_ChromeDriver := RegexReplace(Version_ChromeDriver, "[^\d]*([\.\d]*).*", "$1")

    ;~ DebugWindow("Version Chrome:"  Version_Chrome "`nVersion Chromedriver:" Version_Chromedriver,Clear:=0,LineBreak:=1)

    ; if (mode != "auto_check")
    ; msgbox, ChromeDriver version = %Version_ChromeDriver%
    ; Check if versions are equal

    if InStr(Version_Chromedriver, Version_Chrome){
        if (mode = "auto_check")
        {
            return
        }

        MsgBox,68,Testing,Current Chromedriver is same as Chromeversion.`nDo you still want to download?
        IfMsgBox, No
        {
            Return
        }
    }

    ; Find the matching Chromedriver
    oHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    oHTTP.Open("GET", "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" Version_Chrome, true)
    oHTTP.Send()
    oHTTP.WaitForResponse()
    Version_Chromedriver := oHTTP.ResponseText

    ;~ DebugWindow("The latest release of Chromedriver is:" Version_ChromeDriver,Clear:=0,LineBreak:=1)

    if InStr(Version_Chromedriver, "NoSuchKey"){
        MsgBox,16,Testing,Error`nVersion_Chromedriver
        return
    }

    ; Download the Chromedriver
    Url_ChromeDriver := "https://chromedriver.storage.googleapis.com/" Version_Chromedriver "/chromedriver_win32.zip"
    URLDownloadToFile, %Url_ChromeDriver%, %A_ScriptDir%/chromedriver_win32.zip

    ; Unzip Chromedriver_win32.zip
    fso := ComObjCreate("Scripting.FileSystemObject")
    AppObj := ComObjCreate("Shell.Application")
    FolderObj := AppObj.Namespace(A_ScriptDir "\chromedriver_win32.zip")

    ; msgbox, Pasting Folder

    FilesToZip = A_ScriptDir\chromedriver_win32.zip
    ; FilesToZip = D:\Projects\AHK\_Temp\Test\*.ahk  ;Example of wildcards to compress
    ; FilesToZip := A_ScriptFullPath   ;Example of file to compress
    sZip := A_ScriptDir "\chromedriver_win32.zip"
    sUnz = C:\Users\%A_UserName%\AppData\Local\SeleniumBasic\

    ; Zip(FilesToZip,sZip)
    Sleep, 500
    Unz(sZip,sUnz)
    msgbox, Updated ChromeDriver successfully.

return
}

RunHide(Command) {
    dhw := A_DetectHiddenWindows
    DetectHiddenWindows, On
    Run, %ComSpec%,, Hide, cPid
    WinWait, ahk_pid %cPid%
    DetectHiddenWindows, %dhw%
    DllCall("AttachConsole", "uint", cPid)

    Shell := ComObjCreate("WScript.Shell")
    Exec := Shell.Exec(Command)
    Result := Exec.StdOut.ReadAll()

    DllCall("FreeConsole")
    Process, Close, %cPid%
Return Result
}

;; ----------- 	THE FUNCTIONS   -------------------------------------
Zip(FilesToZip,sZip)
{
    If Not FileExist(sZip)
        CreateZipFile(sZip)
    psh := ComObjCreate( "Shell.Application" )
    pzip := psh.Namespace( sZip )
    if InStr(FileExist(FilesToZip), "D")
        FilesToZip .= SubStr(FilesToZip,0)="\" ? "*.*" : "\*.*"
    loop,%FilesToZip%,1
    {
        zipped++
        ToolTip Zipping %A_LoopFileName% ..
        pzip.CopyHere( A_LoopFileLongPath, 4|16 )
        Loop
        {
            msgbox, % "Done :" done "`nzipped :" zipped
            done := pzip.items().count
            if done = %zipped%
                break
        }
        done := -1
    }
    ToolTip
}

CreateZipFile(sZip)
{
    Header1 := "PK" . Chr(5) . Chr(6)
    VarSetCapacity(Header2, 18, 0)
    file := FileOpen(sZip,"w")
    file.Write(Header1)
    file.RawWrite(Header2,18)
    file.close()
}

Unz(sZip, sUnz)
{
    fso := ComObjCreate("Scripting.FileSystemObject")
    If Not fso.FolderExists(sUnz) ;http://www.autohotkey.com/forum/viewtopic.php?p=402574
        fso.CreateFolder(sUnz)
    psh := ComObjCreate("Shell.Application")
    zippedItems := psh.Namespace(sZip).items().count
    psh.Namespace(sUnz).CopyHere( psh.Namespace( sZip ).items, 4|16 )
    Loop {
        sleep 50
        unzippedItems := psh.Namespace(sUnz).items().count
        ToolTip Unzipping in progress..
        ;    msgbox, % "Zipped items :" zippedItems "`nUnzipped items :" unzippedItems
        ;    IfEqual,zippedItems,%unzippedItems%
        if (unzippedItems) 
            break
    }
    ToolTip
}
;; ----------- 	END FUNCTIONS   -------------------------------------

