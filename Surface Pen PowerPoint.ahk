#SingleInstance force ; allows for only one instance and the new one overwrites the old one


; Surface Pen with single button
; Top button single click is Win+F20 (#F20)
; Top button double click is Win+F19 (#F19
; Top button long press is Win+F18 (#F18)

; Script only responds when PowerPoint presentation is running, so Surface Pen works as used outside PowerPoint
; Shortcut to place exe file in startup: Win+R, "shell:startup"

#If WinActive("ahk_class screenClass")
#F20::send {right}
#F19::send {left}
#F18::send {b}