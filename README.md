<div align="center">

## control subclassing switchboard


</div>

### Description

The Switchboard:A method for handling subclassing in ActiveX controls

f you develop ActiveX controls and intend to subclass or hook a window, you'll very quickly

discover a problem when you attempt to site multiple instances of your control. The subclassing,

which worked fine with a single instance of your control, now no longer works and is, in fact, most

likely is causing a GPF.

Why is this happening? The AddressOf operator requires you to place the callback routine in a

module. This module is shared between all instances of your control and the variables and subroutines

that the module provide are not unique to each instance. The easiest way to visualize the problem is

to imagine a shared phoneline (or a partyline as we hicks call it) where multiple parties are trying to

dial a number, talk, and hangup, all at the same time. What's needed is an operator, a routine that

controls the dialing (hooking), the talking (the callback routine), and who routes information to the

instance of the control that requested it.

The Switchboard subroutine (see below) and it's supporting code provides a method for subclassing

from multiple instances of your ActiveX control. It is not memory intensive, nor is it slow. It's biggest

weakness is that it is hardcoded to intercept particular messages (in this case, WM_SIZE, to trap

resize events) and will require some minor modification on your part to use.
 
### More Info
 
You will find references to myUC in the code below. Replace each instance of this with a reference

to your user control. It is very important that your code detect and respond to a subclassed window when it either closes

(WM_CLOSE) or is destroyed (WM_DESTROY). When this message is received, you should

immediately unhook the window in question. The example code provided here does this, but knowing

why it does it will hopefully save you some grief.

Code Starts

Here

Because this codes hooks into the windows messaging system, you should not use the IDE's STOP

button to terminate the execution of your code. Closing the form normally is mandatory. Debugging

will become difficult once you have subclassed a window, so I recommend adding instancing support

after the bulk of your programming work has been completed. As with any serious API

programming tasks, you should save your project before execution.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tim's VB 5 tips and tricks](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-s-vb-5-tips-and-tricks.md)
**Level**          |Unknown
**User Rating**    |4.2 (165 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-s-vb-5-tips-and-tricks-control-subclassing-switchboard__1-447/archive/master.zip)

### API Declarations

```

Public Const WM_SIZE = &H5
     Public Const GWL_WNDPROC = (-4&)
     Public Const GWL_USERDATA = (-21&)
     Public Const WM_CLOSE = &H10
     Public Const MIN_INSTANCES = 1
     Public Const MAX_INSTANCES = 256
     Type Instances
       in_use As Boolean    'This instance is alive
       ClassAddr As Long    'Pointer to self
       hwnd As Long      'hWnd being hooked
       PrevWndProc As Long   'Stored for unhooking
     End Type
     'Hooking Related Declares
     Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal _
         hwnd As Long, ByVal nIndex As Long)
     Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" _
         (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
         ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long)
     Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" _
         (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
     Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (pDest As Any, pSource As Any, ByVal ByteLen As Long)
     Global Instances(MIN_INSTANCES To MAX_INSTANCES) As Instances
```


### Source Code

```
' Place this code in the General Declarations area
     Dim m_MyInstance as Integer
' Place this block of code in the user control's
     ' INITIALIZE event
       Dim Instance_Scan As Integer
       For Instance_Scan = MIN_INSTANCES To MAX_INSTANCES
         If Instances(Instance_Scan).in_use = False Then
           m_MyInstance = Instance_Scan
           Instances(Instance_Scan).in_use = True
           Instances(Instance_Scan).ClassAddr = ObjPtr(Me)
           Exit For
         End If
       Next Instance_Scan
     ' Note the Friend keyword.
     ' If you plan on modifying wMsg, pass it ByRef...
     Friend Sub ParentResized(ByVal wMsg As Long)
       Static ParentWidth As Long
       Static ParentHeight As Long
       If wMsg = WM_CLOSE Then UnhookParent
       If ParentWidth <> Usercontrol.Parent.Width Or _
         ParentHeight <> Usercontrol.Parent.Height Then
         Debug.Print m_MyInstance & ": Resize event"
       End If
       ParentWidth = TrueParentWidth
       ParentHeight = TrueParentHeight
     End Sub
Public Function SwitchBoard(ByVal hwnd As Long, ByVal MSG As Long, _
         ByVal wParam As Long, ByVal lParam As Long) As Long
       Dim instance_check As Integer
       Dim cMyUC As MyUC
       Dim PrevWndProc As Long
       'Do this early as we may unhook
       PrevWndProc = Is_Hooked(hwnd)
       If MSG = WM_SIZE Or MSG = WM_CLOSE Then
         For instance_check = MIN_INSTANCES To MAX_INSTANCES
           If Instances(instance_check).hwnd = hwnd Then
             On Error Resume Next
             CopyMemory cMyUC, Instances(instance_check).ClassAddr, 4
             cMyUC.ParentResized MSG
             CopyMemory cMyUC, 0&, 4
           End If
         Next instance_check
       End If
       SwitchBoard = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)
     End Function
     'Hooks a window or acts as if it does if the window is
     'already hooked by a previous instance of myUC.
     Public Sub Hook_Window(ByVal hwnd As Long, ByVal instance_ndx As Integer)
       Instances(instance_ndx).PrevWndProc = Is_Hooked(hwnd)
       If Instances(instance_ndx).PrevWndProc = 0& Then
         Instances(instance_ndx).PrevWndProc = SetWindowLong(hwnd, _
           GWL_WNDPROC, AddressOf SwitchBoard)
       End If
       Instances(instance_ndx).hwnd = hwnd
     End Sub
     ' Unhooks only if no other instances need the hWnd
     Public Sub UnHookWindow(ByVal instance_ndx As Integer)
       If TimesHooked(Instances(instance_ndx).hwnd) = 1 Then
         SetWindowLong Instances(instance_ndx).hwnd, GWL_WNDPROC, _
           Instances(instance_ndx).PrevWndProc
       End If
       Instances(instance_ndx).hwnd = 0&
     End Sub
     'Determine if we have already hooked a window,
     'and returns the PrevWndProc if true, 0& if false
     Private Function Is_Hooked(ByVal hwnd As Long) As Long
       Dim ndx As Integer
       Is_Hooked = 0&
       For ndx = MIN_INSTANCES To MAX_INSTANCES
         If Instances(ndx).hwnd = hwnd Then
           Is_Hooked = Instances(ndx).PrevWndProc
           Exit For
         End If
       Next ndx
     End Function
     'Returns a count of the number of times a given
     'window has been hooked by instances of myUC.
     Private Function TimesHooked(ByVal hwnd As Long) As Long
       Dim ndx As Integer
       Dim cnt As Integer
       For ndx = MIN_INSTANCES To MAX_INSTANCES
         If Instances(ndx).hwnd = hwnd Then
           cnt = cnt + 1
         End If
       Next ndx
       TimesHooked = cnt
     End Function
```

