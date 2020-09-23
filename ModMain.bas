Attribute VB_Name = "ModMain"
'This module contains all the important momemtum and forces stuff
'You may use any part of the code here in your own program
'Contact me at nirpaudyal@hotmail.com if you don't understand any part
'Include me in the about box if you felt that i have helped you!
Const Gravity = -1.8
Const Mu = 0.01      'coefficent of friction
'this structure holds the information about the balls
Type ObjectB
    Radius As Single
    X As Single     'the current X position
    Y As Single     'the current Y position
    Vx As Single    'X velocity
    Vy As Single    'Y velocity
    Mass As Single
End Type

Type PointSng
    X As Single
    Y As Single
End Type

'This sub deals with collision dection and bouncing of balls
'I have assumed that the balls all have the same mass
Public Sub ChangeVelocities(a As ObjectB, b As ObjectB)
    Dim X1 As Single, Y1 As Single
    Dim X2 As Single, Y2 As Single
    X1 = a.X
    Y1 = a.Y
    X2 = b.X
    Y2 = b.Y
    'Get the distance between the two balls
    dis = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
    'check to see if a collision has occured
    If dis > a.Radius + b.Radius Then Exit Sub
    'if collision occurs then seperate the balls
    SeperateBalls a, b
    'get the angle between the positions of the balls
    angle = Atn((Y2 - Y1) / (X2 - X1))
    
    hX1 = a.Vx
    hY1 = a.Vy
    hX2 = b.Vx
    hY2 = b.Vy
    'resolve the velocitis such that they are along the line of collision
    X1 = hX1 * Cos(-angle) - hY1 * Sin(-angle)
    Y1 = hX1 * Sin(-angle) + hY1 * Cos(-angle)
    X2 = hX2 * Cos(-angle) - hY2 * Sin(-angle)
    Y2 = hX2 * Sin(-angle) + hY2 * Cos(-angle)
    'swap the horizontal components of the velocities
    '(do any momemtum calculations here)
    hX1 = (X1 * (a.Mass - b.Mass) + (X2 * 2 * b.Mass)) / (a.Mass + b.Mass)
    hX2 = ((X1 * 2 * a.Mass) + X2 * (a.Mass - b.Mass)) / (a.Mass + b.Mass)
    'keep the vertical component the same
    hY1 = Y1
    hY2 = Y2
    'resolve back the velocities to their normal coordinates
    X1 = hX1 * Cos(angle) - hY1 * Sin(angle)
    Y1 = hX1 * Sin(angle) + hY1 * Cos(angle)
    X2 = hX2 * Cos(angle) - hY2 * Sin(angle)
    Y2 = hX2 * Sin(angle) + hY2 * Cos(angle)
    'set the velocities of the ball
    a.Vx = X1
    a.Vy = Y1
    b.Vx = X2
    b.Vy = Y2
End Sub
Public Sub SeperateBalls(a As ObjectB, b As ObjectB)
    'reset the position of the balls so that they dont overlap
    'this process is achieved using similar triangles
    dx = (b.X - a.X)
    dy = (b.Y - a.Y)
    l = Sqr(dx * dx + dy * dy)
    G = (a.Radius + b.Radius) - l
    DeltaX = (G / l) * dx
    DeltaY = (G / l) * dy
    b.X = b.X + DeltaX
    b.Y = b.Y + DeltaY
End Sub
Public Sub HandleFriction(P As ObjectB)
    'get the speed of the ball
    V = Sqr(P.Vx * P.Vx + P.Vy * P.Vy)
    'friction doesn't act while ball is not in motion
    If V < 0 Then Exit Sub
    'if speed is really low then set it to zero
    If V < 0.1 Then
        P.Vx = 0
        P.Vy = 0
        Exit Sub
    End If
    Dim fx As Single
    Dim fy As Single
    'calculate the friction
    Friction = Mu * P.Mass * Abs(Gravity)
    If P.Vx = 0 Then ang = 0 Else ang = Atn(P.Vy / P.Vx)
    'get the components of frictions in the two directions
    fx = Abs(Friction * Cos(ang))
    fy = Abs(Friction * Sin(ang))
    'ensure that the friction is opposing the direction of motion
    If P.Vx > 0 Then fx = -fx
    If P.Vy > 0 Then fy = -fy
    'apply the force
    ApplyForce P, fx, fy, 0.1
End Sub
Sub ApplyForce(P As ObjectB, ForceX As Single, ForceY As Single, Time_Of_Force As Single)
    'Use F= (mv-mu)/t to find v, the new velocity of the ball once the force is applied
    P.Vx = P.Vx + (ForceX * Time_Of_Force / P.Mass)
    P.Vy = P.Vy + (ForceY * Time_Of_Force / P.Mass)
End Sub
