VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "GlassCrack"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form2"
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   1800
      Left            =   720
      Top             =   180
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Animation phases
Private Const PHASE_CRACK = 1
Private Const PHASE_SHATTER = 2
Private Const PHASE_PAUSE = 3    ' beat between waves; demon visible through central hole
Private Const PHASE_SHATTER2 = 4 ' rest of the screen drops away
Private Const PHASE_FINAL = 5
Private Const PHASE_DONE = 6

Private m_phase As Long
Private m_frame As Long
Private m_phaseStart As Long

Private m_scrW As Long, m_scrH As Long
Private m_cx As Long, m_cy As Long

' Desktop screenshot in a memory DC
Private m_backHdc As Long
Private m_backBmp As Long
Private m_backOldObj As Long

' Cracks
Private Type CrackLine
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type
Private m_cracks() As CrackLine
Private m_nCracks As Long

' --- The new shard model: each shard is a triangle with a home polygon
'     on the desktop, plus current displacement/rotation as it falls.
Private Type Shard
    ' Home polygon (3 points, in screen coords, defining where this shard
    ' originally sat on the desktop). p0 is the impact-center vertex.
    p0x As Single: p0y As Single
    p1x As Single: p1y As Single
    p2x As Single: p2y As Single
    ' Centroid of home polygon, used as rotation pivot
    cxh As Single: cyh As Single
    ' Distance from impact center (used to order detachment waves)
    distFromImpact As Single
    ' Detachment frame (relative to PHASE_SHATTER start) and falling state
    detachFrame As Long
    detached As Boolean
    ' Current displacement from home position
    dx As Single: dy As Single
    vx As Single: vy As Single
    rot As Single  ' radians
    vrot As Single
    wave As Long   ' 1 = initial impact, 2 = secondary collapse
End Type
Private m_shards() As Shard
Private m_nShards As Long

' Demon
Private m_demonImg As Long
Private m_demonW As Long, m_demonH As Long

' Phase durations (frames @ ~60fps)
Private Const FR_CRACK = 25
Private Const FR_SHATTER = 60   ' wave 1: inner shards (rings 0-1) detach and fall
Private Const FR_PAUSE = 25     ' beat between waves (~0.4s) — demon glimpse
Private Const FR_SHATTER2 = 90  ' wave 2: outer shards (rings 2-4) collapse
Private Const FR_FINAL = 50     ' demon zooms to full and roars

Private Sub Form_Load()
    On Error GoTo Trap

    Dim exeDir As String
    exeDir = App.path
    If Right$(exeDir, 1) <> "\" Then exeDir = exeDir & "\"

    GdipInit
    If FileThere(exeDir & "demon.png") Then
        m_demonImg = LoadPNG(exeDir & "demon.png")
        If m_demonImg <> 0 Then
            GdipGetImageWidth m_demonImg, m_demonW
            GdipGetImageHeight m_demonImg, m_demonH
        End If
    End If

    m_scrW = GetSystemMetrics(SM_CXSCREEN)
    m_scrH = GetSystemMetrics(SM_CYSCREEN)
    m_cx = m_scrW \ 2
    m_cy = m_scrH \ 2

    ' Capture desktop into back DC
    Dim screenDC As Long
    screenDC = GetDC(0&)
    m_backHdc = CreateCompatibleDC(screenDC)
    m_backBmp = CreateCompatibleBitmap(screenDC, m_scrW, m_scrH)
    m_backOldObj = SelectObject(m_backHdc, m_backBmp)
    BitBlt m_backHdc, 0, 0, m_scrW, m_scrH, screenDC, 0, 0, SRCCOPY
    ReleaseDC 0&, screenDC

    ' Show identical desktop first
    BitBlt Me.hDC, 0, 0, m_scrW, m_scrH, m_backHdc, 0, 0, SRCCOPY
    Me.Refresh

    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW

    BuildCracks
    BuildShards

    If FileThere(exeDir & "glass_crack.wav") Then
        sndPlaySound exeDir & "glass_crack.wav", SND_ASYNC Or SND_FILENAME Or SND_NODEFAULT
    End If

    m_phase = PHASE_CRACK
    m_frame = 0
    m_phaseStart = 0
    tmr.Enabled = True
    Exit Sub

Trap:
    ' If anything blew up, drop topmost so error dialog is visible and reachable.
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    tmr.Enabled = False
    MsgBox "Form_Load error #" & Err.Number & ": " & Err.Description, vbCritical, "GlassCrack"
    Unload Me
End Sub

' ---- Build a radial shatter pattern: rings of triangles around impact center.
'      Inner ring shards are smaller (more detail near impact), outer larger.
Private Sub BuildShards()
    Dim ringRadii(0 To 4) As Single
    Dim ringSegs(0 To 4) As Long
    Dim i As Long, j As Long, idx As Long
    Dim a1 As Single, a2 As Single, jitter As Single

    ringRadii(0) = 60
    ringRadii(1) = 140
    ringRadii(2) = 240
    ringRadii(3) = 360
    ringRadii(4) = 520  ' last ring extends past viewport edges - covers everywhere
    ringSegs(0) = 8
    ringSegs(1) = 12
    ringSegs(2) = 14
    ringSegs(3) = 16
    ringSegs(4) = 18

    ' Total shards = sum of segs across rings (each segment = 2 triangles, but
    ' we'll just use 1 triangular shard per segment for cleaner geometry)
    Dim total As Long
    For i = 0 To 4
        total = total + ringSegs(i)
    Next i
    ReDim m_shards(1 To total) As Shard
    m_nShards = 0

    Randomize Timer

    For i = 0 To 4
        Dim segCount As Long
        segCount = ringSegs(i)
        Dim rIn As Single, rOut As Single
        If i = 0 Then
            rIn = 0
        Else
            rIn = ringRadii(i - 1)
        End If
        rOut = ringRadii(i)

        For j = 0 To segCount - 1
            a1 = (j / segCount) * 6.2831853 + (Rnd - 0.5) * 0.15
            a2 = ((j + 1) / segCount) * 6.2831853 + (Rnd - 0.5) * 0.15

            idx = m_nShards + 1
            m_nShards = idx

            If i = 0 Then
                ' Innermost ring: triangle with apex at center, base at ring edge
                jitter = (Rnd - 0.5) * 15
                m_shards(idx).p0x = m_cx
                m_shards(idx).p0y = m_cy
                m_shards(idx).p1x = m_cx + Cos(a1) * (rOut + jitter)
                m_shards(idx).p1y = m_cy + Sin(a1) * (rOut + jitter)
                m_shards(idx).p2x = m_cx + Cos(a2) * (rOut + jitter)
                m_shards(idx).p2y = m_cy + Sin(a2) * (rOut + jitter)
            Else
                ' Outer rings: triangle alternates orientation so they tile.
                ' Two vertices on outer arc, one on inner arc — irregular for shatter look.
                jitter = (Rnd - 0.5) * 20
                Dim aMid As Single
                aMid = (a1 + a2) * 0.5
                If (j Mod 2) = 0 Then
                    m_shards(idx).p0x = m_cx + Cos(aMid) * (rIn + jitter)
                    m_shards(idx).p0y = m_cy + Sin(aMid) * (rIn + jitter)
                    m_shards(idx).p1x = m_cx + Cos(a1) * (rOut + jitter)
                    m_shards(idx).p1y = m_cy + Sin(a1) * (rOut + jitter)
                    m_shards(idx).p2x = m_cx + Cos(a2) * (rOut + jitter)
                    m_shards(idx).p2y = m_cy + Sin(a2) * (rOut + jitter)
                Else
                    m_shards(idx).p0x = m_cx + Cos(aMid) * (rOut + jitter)
                    m_shards(idx).p0y = m_cy + Sin(aMid) * (rOut + jitter)
                    m_shards(idx).p1x = m_cx + Cos(a1) * (rIn + jitter)
                    m_shards(idx).p1y = m_cy + Sin(a1) * (rIn + jitter)
                    m_shards(idx).p2x = m_cx + Cos(a2) * (rIn + jitter)
                    m_shards(idx).p2y = m_cy + Sin(a2) * (rIn + jitter)
                End If
            End If

            ' Centroid + distance
            m_shards(idx).cxh = (m_shards(idx).p0x + m_shards(idx).p1x + m_shards(idx).p2x) / 3
            m_shards(idx).cyh = (m_shards(idx).p0y + m_shards(idx).p1y + m_shards(idx).p2y) / 3
            Dim ddx As Single, ddy As Single
            ddx = m_shards(idx).cxh - m_cx
            ddy = m_shards(idx).cyh - m_cy
            m_shards(idx).distFromImpact = Sqr(ddx * ddx + ddy * ddy)

            ' Two-wave system:
            '   Wave 1 = rings 0-1 (closest 20 shards). Detach within first
            '     ~20 frames of PHASE_SHATTER for a punchy initial spray.
            '   Wave 2 = rings 2-4 (outer 48 shards). Detach in first ~15
            '     frames of PHASE_SHATTER2, all roughly together — the
            '     surrounding glass giving way at the second crack sound.
            ' detachFrame is interpreted relative to the shard's wave phase.
            If i <= 1 Then
                m_shards(idx).wave = 1
                m_shards(idx).detachFrame = CLng(i * 6 + Rnd * 6)  ' 0..12 for ring 0, 6..18 for ring 1
            Else
                m_shards(idx).wave = 2
                m_shards(idx).detachFrame = CLng(Rnd * 15)  ' all detach in first 15 frames of wave 2
            End If
            m_shards(idx).detached = False

            ' Falling velocity differs by wave.
            ' Wave 1: dramatic outward spray (small shards exploding out from impact).
            ' Wave 2: more "slide and fall" — larger pieces giving way under gravity.
            Dim outAng As Single
            outAng = Atan2Approx(ddy, ddx)
            If m_shards(idx).wave = 1 Then
                m_shards(idx).vx = Cos(outAng) * (1 + Rnd * 2.5)
                m_shards(idx).vy = Sin(outAng) * (0.5 + Rnd * 1.5) + 0.5 + Rnd * 1.2
                m_shards(idx).vrot = (Rnd - 0.5) * 0.12
            Else
                m_shards(idx).vx = Cos(outAng) * (0.3 + Rnd * 0.8)
                m_shards(idx).vy = 0.3 + Rnd * 0.8
                m_shards(idx).vrot = (Rnd - 0.5) * 0.06
            End If
            m_shards(idx).rot = 0
            m_shards(idx).dx = 0
            m_shards(idx).dy = 0
        Next j
    Next i
End Sub

' VB6 has no Atan2; quick approximation good enough for direction vectors.
Private Function Atan2Approx(y As Single, x As Single) As Single
    If x > 0 Then
        Atan2Approx = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atan2Approx = Atn(y / x) + 3.14159265
        Else
            Atan2Approx = Atn(y / x) - 3.14159265
        End If
    Else
        If y > 0 Then
            Atan2Approx = 1.5707963
        ElseIf y < 0 Then
            Atan2Approx = -1.5707963
        Else
            Atan2Approx = 0
        End If
    End If
End Function

Private Sub BuildCracks()
    Dim n As Long, i As Long, j As Long
    n = 14
    ReDim m_cracks(1 To n * 6) As CrackLine
    m_nCracks = 0
    Randomize Timer
    Dim angle As Single, length As Long
    Dim px As Long, py As Long, nx As Long, ny As Long
    Dim segCount As Long, segLen As Long, jitter As Single
    For i = 1 To n
        angle = (i - 1) * (6.2831853 / n) + (Rnd - 0.5) * 0.3
        length = 180 + Int(Rnd * 220)
        segCount = 3 + Int(Rnd * 3)
        segLen = length \ segCount
        px = m_cx
        py = m_cy
        For j = 1 To segCount
            jitter = (Rnd - 0.5) * 0.6
            nx = px + segLen * Cos(angle + jitter)
            ny = py + segLen * Sin(angle + jitter)
            m_nCracks = m_nCracks + 1
            m_cracks(m_nCracks).x1 = px
            m_cracks(m_nCracks).y1 = py
            m_cracks(m_nCracks).x2 = nx
            m_cracks(m_nCracks).y2 = ny
            px = nx
            py = ny
        Next j
    Next i
End Sub

' ===== Main animation tick =====
Private Sub tmr_Timer()
    On Error GoTo Trap
    m_frame = m_frame + 1
    Dim loc As Long
    loc = m_frame - m_phaseStart

    Select Case m_phase
        Case PHASE_CRACK
            BitBlt Me.hDC, 0, 0, m_scrW, m_scrH, m_backHdc, 0, 0, SRCCOPY
            DrawCracks loc, FR_CRACK
            If loc >= FR_CRACK Then NextPhase PHASE_SHATTER

        Case PHASE_SHATTER
            ' Demon ramps from 0 to 0.4 across wave 1
            RenderShatterFrame loc, 1, (loc / FR_SHATTER) * 0.4
            If loc >= FR_SHATTER Then NextPhase PHASE_PAUSE

        Case PHASE_PAUSE
            ' Wave-1 shards still falling; no new detachments (activeWave=0 ⇒ nothing matches).
            ' Demon ramps 0.4 -> 0.6 during pause — clearly visible through central hole.
            Dim pauseProg As Single
            pauseProg = loc / FR_PAUSE
            If pauseProg > 1 Then pauseProg = 1
            RenderShatterFrame loc, 0, 0.4 + pauseProg * 0.2
            If loc >= FR_PAUSE Then
                PlaySecondCrack
                NextPhase PHASE_SHATTER2
            End If

        Case PHASE_SHATTER2
            ' Wave 2: outer shards detach and fall. Demon ramps 0.6 -> 0.85.
            Dim w2Prog As Single
            w2Prog = loc / FR_SHATTER2
            If w2Prog > 1 Then w2Prog = 1
            RenderShatterFrame loc, 2, 0.6 + w2Prog * 0.25
            If loc >= FR_SHATTER2 Then NextPhase PHASE_FINAL

        Case PHASE_FINAL
            RenderFinalFrame loc
            If loc >= FR_FINAL Then m_phase = PHASE_DONE

        Case PHASE_DONE
            RenderFinalFrame FR_FINAL
            tmrClose.Enabled = True
    End Select

    Me.Refresh
    Exit Sub

Trap:
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    tmr.Enabled = False
    MsgBox "Timer error #" & Err.Number & ": " & Err.Description, vbCritical, "GlassCrack"
End Sub

Private Sub NextPhase(ByVal p As Long)
    m_phase = p
    m_phaseStart = m_frame
End Sub

' ===== Shatter phase rendering =====
'
' Each frame:
'  1. Update shard physics (detach when frame reached, then fall + rotate)
'  2. Fill screen black (this is the "void behind the glass")
'  3. Draw the lurking demon, dim and small, in the center
'  4. Build a clip region = full screen MINUS all detached shards' home polygons.
'     Blit the desktop bitmap with that clip → desktop appears with holes where
'     shards have fallen away.
'  5. Draw cracks on top of remaining desktop
'  6. For each detached shard, set clip to its current (translated+rotated)
'     polygon and blit the desktop bitmap with an offset that maps from the
'     shard's home position to its current position. The shard carries its
'     wallpaper texture as it falls.
'
Private Sub RenderShatterFrame(ByVal localFr As Long, ByVal activeWave As Long, ByVal demonProgress As Single)
    UpdateShardPhysics localFr, activeWave

    ' --- Layer 1: black void ---
    PatBlt Me.hDC, 0, 0, m_scrW, m_scrH, BLACKNESS

    ' --- Layer 2: demon lurking, growing across the whole shatter sequence ---
    DrawLurkingDemon demonProgress

    ' --- Layer 3: remaining desktop (everywhere except where shards fell from)
    Dim clipRgn As Long
    clipRgn = BuildDesktopClipRegion()
    SelectClipRgn Me.hDC, clipRgn
    BitBlt Me.hDC, 0, 0, m_scrW, m_scrH, m_backHdc, 0, 0, SRCCOPY
    SelectClipRgn Me.hDC, 0  ' release clip
    DeleteObject clipRgn

    ' --- Layer 4: cracks drawn only on intact desktop (use same clip) ---
    clipRgn = BuildDesktopClipRegion()
    SelectClipRgn Me.hDC, clipRgn
    DrawCracks FR_CRACK, FR_CRACK
    SelectClipRgn Me.hDC, 0
    DeleteObject clipRgn

    ' --- Layer 5: falling shards (textured with desktop pixels) ---
    Dim i As Long
    For i = 1 To m_nShards
        If m_shards(i).detached Then DrawFallingShard i
    Next i
End Sub

Private Sub UpdateShardPhysics(ByVal localFr As Long, ByVal activeWave As Long)
    Dim i As Long
    For i = 1 To m_nShards
        If Not m_shards(i).detached Then
            ' Only detach shards belonging to the currently active wave.
            If m_shards(i).wave = activeWave Then
                If localFr >= m_shards(i).detachFrame Then
                    m_shards(i).detached = True
                End If
            End If
        End If
        ' Physics applies to ALL already-detached shards regardless of wave,
        ' so wave-1 pieces keep falling during pause and wave 2.
        If m_shards(i).detached Then
            m_shards(i).vy = m_shards(i).vy + 0.55  ' gravity
            m_shards(i).dx = m_shards(i).dx + m_shards(i).vx
            m_shards(i).dy = m_shards(i).dy + m_shards(i).vy
            m_shards(i).rot = m_shards(i).rot + m_shards(i).vrot
        End If
    Next i
End Sub

' Build the region representing intact desktop: full screen minus union of all
' detached shards' home polygons.
Private Function BuildDesktopClipRegion() As Long
    Dim full As Long, holes As Long, tmp As Long, polyRgn As Long
    full = CreateRectRgn(0, 0, m_scrW, m_scrH)
    holes = CreateRectRgn(0, 0, 0, 0)  ' empty

    Dim pts(0 To 2) As POINTAPI
    Dim i As Long
    For i = 1 To m_nShards
        If m_shards(i).detached Then
            pts(0).x = CLng(m_shards(i).p0x)
            pts(0).y = CLng(m_shards(i).p0y)
            pts(1).x = CLng(m_shards(i).p1x)
            pts(1).y = CLng(m_shards(i).p1y)
            pts(2).x = CLng(m_shards(i).p2x)
            pts(2).y = CLng(m_shards(i).p2y)
            polyRgn = CreatePolygonRgn(pts(0), 3, ALTERNATE)
            If polyRgn <> 0 Then
                tmp = CreateRectRgn(0, 0, 0, 0)
                CombineRgn tmp, holes, polyRgn, RGN_OR
                DeleteObject holes
                holes = tmp
                DeleteObject polyRgn
            End If
        End If
    Next i

    Dim result As Long
    result = CreateRectRgn(0, 0, 0, 0)
    CombineRgn result, full, holes, RGN_DIFF
    DeleteObject full
    DeleteObject holes
    BuildDesktopClipRegion = result
End Function

' Draw a single falling shard. We rotate the home polygon's three vertices
' around its centroid, then translate by (dx, dy). The desktop bitmap is
' blitted with an offset chosen so the shard's centroid still samples its
' original desktop pixels — the texture follows the polygon. (Rotation of the
' actual pixels isn't applied; only the shape rotates. For small rotations and
' fast motion this reads correctly to the eye.)
Private Sub DrawFallingShard(ByVal idx As Long)
    Dim cs As Single, sn As Single
    cs = Cos(m_shards(idx).rot)
    sn = Sin(m_shards(idx).rot)
    Dim cxh As Single, cyh As Single
    cxh = m_shards(idx).cxh
    cyh = m_shards(idx).cyh

    Dim pts(0 To 2) As POINTAPI
    Dim hx As Single, hy As Single, rx As Single, ry As Single

    ' p0
    hx = m_shards(idx).p0x - cxh: hy = m_shards(idx).p0y - cyh
    rx = hx * cs - hy * sn: ry = hx * sn + hy * cs
    pts(0).x = CLng(cxh + rx + m_shards(idx).dx)
    pts(0).y = CLng(cyh + ry + m_shards(idx).dy)
    ' p1
    hx = m_shards(idx).p1x - cxh: hy = m_shards(idx).p1y - cyh
    rx = hx * cs - hy * sn: ry = hx * sn + hy * cs
    pts(1).x = CLng(cxh + rx + m_shards(idx).dx)
    pts(1).y = CLng(cyh + ry + m_shards(idx).dy)
    ' p2
    hx = m_shards(idx).p2x - cxh: hy = m_shards(idx).p2y - cyh
    rx = hx * cs - hy * sn: ry = hx * sn + hy * cs
    pts(2).x = CLng(cxh + rx + m_shards(idx).dx)
    pts(2).y = CLng(cyh + ry + m_shards(idx).dy)

    ' Off-screen quick reject
    If pts(0).y > m_scrH + 100 And pts(1).y > m_scrH + 100 And pts(2).y > m_scrH + 100 Then Exit Sub

    Dim polyRgn As Long
    polyRgn = CreatePolygonRgn(pts(0), 3, ALTERNATE)
    If polyRgn = 0 Then Exit Sub

    SelectClipRgn Me.hDC, polyRgn
    ' BitBlt with negative offset to pull desktop pixels from the home location
    ' to the shard's current location.
    Dim offX As Long, offY As Long
    offX = CLng(m_shards(idx).dx)
    offY = CLng(m_shards(idx).dy)
    BitBlt Me.hDC, offX, offY, m_scrW, m_scrH, m_backHdc, 0, 0, SRCCOPY

    ' Subtle dark edge along the shard outline so pieces read as discrete
    Dim pen As Long, oldPen As Long
    pen = CreatePen(PS_SOLID, 1, RGB(20, 20, 22))
    oldPen = SelectObject(Me.hDC, pen)
    SelectClipRgn Me.hDC, 0  ' draw edge unclipped
    Dim outline(0 To 3) As POINTAPI
    outline(0) = pts(0): outline(1) = pts(1): outline(2) = pts(2): outline(3) = pts(0)
    MoveToEx Me.hDC, outline(0).x, outline(0).y, 0
    LineTo Me.hDC, outline(1).x, outline(1).y
    LineTo Me.hDC, outline(2).x, outline(2).y
    LineTo Me.hDC, outline(0).x, outline(0).y
    SelectObject Me.hDC, oldPen
    DeleteObject pen

    DeleteObject polyRgn
End Sub

' ===== Demon drawing =====

' During shatter phase: small, dim, lurking in the void
Private Sub DrawLurkingDemon(ByVal progress As Single)
    Dim s As Single
    s = 0.15 + progress * 0.35  ' 15% -> 50% of final
    DrawDemonAtScale s, 0.4 + progress * 0.5  ' brightness ramp via... well, just size for now
End Sub

' Final phase: full zoom + scream
Private Sub RenderFinalFrame(ByVal localFr As Long)
    Dim t As Single
    t = localFr / FR_FINAL
    If t > 1 Then t = 1
    Dim eased As Single
    eased = 1 - (1 - t) * (1 - t)

    PatBlt Me.hDC, 0, 0, m_scrW, m_scrH, BLACKNESS

    Dim s As Single
    s = 0.5 + 0.5 * eased  ' 50% -> 100% during final
    DrawDemonAtScale s, 1
    
    If localFr = 1 Then PlayRoar
End Sub

Private Sub DrawDemonAtScale(ByVal scal As Single, ByVal opacityUnused As Single)
    Dim targetSide As Long
    targetSide = CLng(m_scrH * 0.95 * scal)
    Dim curW As Long, curH As Long
    curW = targetSide: curH = targetSide
    Dim dx As Long, dy As Long
    dx = m_cx - curW \ 2
    dy = m_cy - curH \ 2
    If m_demonImg <> 0 Then
        DrawPNG Me.hDC, m_demonImg, dx, dy, curW, curH
    Else
        ' Fallback so something visible appears
        Dim br As Long, oldBr As Long, pn As Long, oldPn As Long
        br = CreateSolidBrush(RGB(140, 15, 15))
        pn = CreatePen(PS_SOLID, 2, RGB(220, 60, 60))
        oldBr = SelectObject(Me.hDC, br)
        oldPn = SelectObject(Me.hDC, pn)
        Dim pts(0 To 3) As POINTAPI
        pts(0).x = dx + curW \ 2: pts(0).y = dy
        pts(1).x = dx + curW: pts(1).y = dy + curH \ 2
        pts(2).x = dx + curW \ 2: pts(2).y = dy + curH
        pts(3).x = dx: pts(3).y = dy + curH \ 2
        Polygon Me.hDC, pts(0), 4
        SelectObject Me.hDC, oldBr
        SelectObject Me.hDC, oldPn
        DeleteObject br
        DeleteObject pn
    End If
End Sub

Private Sub DrawCracks(ByVal localFr As Long, ByVal totalFr As Long)
    Dim t As Single
    t = localFr / totalFr
    If t > 1 Then t = 1
    ' Original layered look (dark halo + bright core) but heavier:
    '   - halo 5px, pure black (was 3px near-black)
    '   - core 1px, bright but not full white (so the black reads stronger)
    Dim penDark As Long, penLight As Long, oldPen As Long
    penDark = CreatePen(PS_SOLID, 5, RGB(0, 0, 0))
    penLight = CreatePen(PS_SOLID, 1, RGB(220, 220, 225))
    Dim i As Long, drawTo As Long
    drawTo = Int(m_nCracks * t)
    If drawTo < 1 Then
        DeleteObject penDark: DeleteObject penLight
        Exit Sub
    End If
    oldPen = SelectObject(Me.hDC, penDark)
    For i = 1 To drawTo
        MoveToEx Me.hDC, m_cracks(i).x1, m_cracks(i).y1, 0
        LineTo Me.hDC, m_cracks(i).x2, m_cracks(i).y2
    Next i
    SelectObject Me.hDC, penLight
    For i = 1 To drawTo
        MoveToEx Me.hDC, m_cracks(i).x1, m_cracks(i).y1, 0
        LineTo Me.hDC, m_cracks(i).x2, m_cracks(i).y2
    Next i
    SelectObject Me.hDC, oldPen
    DeleteObject penDark
    DeleteObject penLight
End Sub

Private Sub PlayRoar()
    Dim p As String
    p = App.path
    If Right$(p, 1) <> "\" Then p = p & "\"
    p = p & "demon_roar.wav"
    If FileThere(p) Then sndPlaySound p, SND_ASYNC Or SND_FILENAME Or SND_NODEFAULT
End Sub

Private Sub PlaySecondCrack()
    Dim p As String
    p = App.path
    If Right$(p, 1) <> "\" Then p = p & "\"
    p = p & "glass_crack.wav"
    If FileThere(p) Then sndPlaySound p, SND_ASYNC Or SND_FILENAME Or SND_NODEFAULT
End Sub

' ===== Exit =====
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub
Private Sub Form_Click()
    Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmr.Enabled = False
    sndPlaySound vbNullString, SND_ASYNC
    SelectClipRgn Me.hDC, 0
    If m_backHdc <> 0 Then
        If m_backOldObj <> 0 Then SelectObject m_backHdc, m_backOldObj
        If m_backBmp <> 0 Then DeleteObject m_backBmp
        DeleteDC m_backHdc
    End If
    If m_demonImg <> 0 Then GdipDisposeImage m_demonImg
    GdipShutdown
End Sub

Private Sub tmrClose_Timer()
    Unload Me
End Sub
