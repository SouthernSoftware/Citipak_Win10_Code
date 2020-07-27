Attribute VB_Name = "Grad"

'Gradient Background Source code - Released into the public domain by John Rogers, June 19, 1996
'
'Requires VB40032.DLL.
'Gradient Background Demonstration program requires COMCTL32.OCX and THREED32.OCX
'
'   This was written in 800x600 mode, so my apologies to those running in 640x480. >:-P
'
'   Quicky destructions: To gradient a form with, say, the blue-to-black gradient found in
'most setup programs, you would put the command
'                            Gradient Me, 0, 0, 255, 1
'into the Resize sub. In the form's properties, turn on AutoRedraw, set the Appearance to Flat
'and you're done! Compile the program and admire your nice gradient! Warning: Due to Windows'
'lousy dithering, this will look absolutely TERRIBLE in anything less than 16-bit (high) color.
'But then again, so do all those setup programs >:-)
'How it works:
'   Pretty simple, really. It just divides the object into 63 sections and fills each one with
'a slightly darker color than the previous one, starting with the given RGB value and ending
'with black. I think that was a run-on, but who cares. It's not like this is documentation.
'For a semi-nifty effect, try commenting one or two of the decrement lines. You'll wind up with
'a two-color gradient. You can also make sideways gradients by swapping a few variables. Knock
'yourself out; this is public domain, which means you can alter it to your heart's content! Have
'fun! Incidentally, the demo program does have a real use: you can use it to design a nicely
'colored background, then write down the syntax. Type it into VB as it is shown, and you'll get
'Your gradient just as it appeared! (If you don't, you most likely ) Like this program?
'Drop me a line at patr@xanadu2.net. Happy shading!
'
Sub Gradient(TheObject As Object, Redval&, Greenval&, Blueval&, TopToBottom As Boolean)
    'TheObject can be any object that supports the Line method (like forms and pictures).
    'Redval, Greenval, and Blueval are the Red, Green, and Blue starting values from 0 to 255.
    'TopToBottom determines whether the gradient will draw down or up.
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$, Stepper%
    Stepper = 255
    'This will create 63 steps in the gradient. This looks smooth on 16-bit and 24-bit color.
    'You can change this, but be careful. You can do some strange-looking stuff with it...
    Step = (TheObject.Height / Stepper)
    'This tells it whether to start on the top or the bottom and adjusts variables accordingly.
    If TopToBottom = True Then FillTop = 0 Else FillTop = TheObject.Height - Step
    FillLeft = 0
    FillRight = TheObject.Width
    FillBottom = FillTop + Step
    'If you changed the number of steps, change the number of reps to match it.
    'If you don't, the gradient will look all funny.
    For Reps = 1 To Stepper
        'This draws the colored bar.
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
        'This decreases the RGB values to darken the color.
        'Lower the value for "squished" gradients. Raise it for incomplete gradients.
        'Also, if you change the number of steps, you will need to change this number.
        Redval = Redval - 1
        Greenval = Greenval - 1
        Blueval = Blueval - 1
        'This prevents the RGB values from becoming negative, which causes a runtime error.
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        'More top or bottom stuff; Moves to next bar.
        If TopToBottom = True Then FillTop = FillBottom Else FillTop = FillTop - Step
        FillBottom = FillTop + Step
    Next
End Sub


