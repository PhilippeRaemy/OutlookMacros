Attribute VB_Name = "COnstants"
Enum WshWindowStyle
    WshWSHideParentActivateChild = 0 'Hides the window and activates another window.
    WshWSActivateChildRestore1st = 1 ' Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
    WshWSActivateChildMinimized = 2 'Activates the window and displays it as a minimized window.
    WshWSActivateChildMaximized = 3 'Activates the window and displays it as a maximized window.
    WshWSActivateChildNoFocus = 4 'Displays a window in its most recent size and position. The active window remains active.
    WshWSActivateChildFocus = 5 ' Activates the window and displays it in its current size and position.
    WshWSMinimizeParent = 6 'Minimizes the specified window and activates the next top-level window in the Z order.
    WshWSDisplayChildMinimized = 7 'Displays the window as a minimized window. The active window remains active.
    WshWSDisplayChild = 8 'Displays the window in its current state. The active window remains active.
    WshWSActivateChildRestore = 9 ' Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    WshWSDefaultShowState = 10 'Sets the show-state based on the state of the program that started the application.
End Enum

