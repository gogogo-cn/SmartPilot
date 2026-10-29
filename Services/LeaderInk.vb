Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Namespace Services

    ''' <summary>
    ''' 用引线墨迹找同一条尺寸的文字框。
    ''' 单头折线：名义尺寸和上下偏差落在水平尾段上。
    ''' 双头：字落在两个箭头之间的那一段上。
    ''' 落在同一段引线上的球标和尺寸收成一条。基准圆仍单独留下。
    ''' </summary>
    Public Class LeaderInk

        Private Class Landing
            Public Horizontal As Boolean
            Public Line As Integer
            Public Start As Integer
            Public Finish As Integer
        End Class

        Public Shared Sub Apply(bmp As Bitmap, cells As List(Of OcrElement))
            If cells Is Nothing OrElse cells.Count = 0 Then Return
            Dim ink = If(bmp Is Nothing, Nothing, LoadInk(bmp))
            If ink IsNot Nothing Then
                Dim claimed As New HashSet(Of OcrElement)()
                Dim seeds = cells.Where(Function(el) IsSeed(el)).OrderByDescending(Function(el) el.BoundingBox.Bottom).ThenBy(Function(el) el.BoundingBox.Left).ToList()
                Dim remove As New HashSet(Of OcrElement)()
                Dim add As New List(Of OcrElement)()

                For Each seed In seeds
                    If claimed.Contains(seed) Then Continue For
                    Dim members = Bind(ink, seed, cells, True)
                    If members Is Nothing Then members = Bind(ink, seed, cells, False)
                    If members Is Nothing OrElse members.Count < 2 Then Continue For
                    Dim fresh = members.Where(Function(el) Not claimed.Contains(el)).ToList()
                    If Not fresh.Contains(seed) OrElse fresh.Count < 2 Then Continue For
                    For Each el In fresh
                        claimed.Add(el)
                        remove.Add(el)
                    Next
                    Dim block = DrawingDimensionGrouper.JoinOnLeader(fresh)
                    add.Add(block)
                    System.Diagnostics.Debug.WriteLine("[引线] " & block.Text)
                Next

                If remove.Count > 0 Then
                    cells.RemoveAll(Function(el) remove.Contains(el))
                    cells.AddRange(add)
                End If
            End If

            ' 球标和尺寸是两个文字框，共用同一条引线。引线墨迹把它们连上才收成一块。
            GroupBySharedLeader(ink, cells)
            ' 尺寸线两端有箭头。球标落在箭头旁，穿过的中心线不算另一条引线。
            AttachAtArrow(ink, cells)
        End Sub

        Private Shared Sub AttachAtArrow(ink As Boolean(,), cells As List(Of OcrElement))
            If ink Is Nothing Then Return
            Dim hosts = cells.Where(Function(el) IsCalloutHost(el) AndAlso Not IsBalloonLike(el)).ToList()
            Dim tags = cells.Where(Function(el) IsArrowTag(el)).ToList()
            If hosts.Count = 0 OrElse tags.Count = 0 Then Return

            Dim ranked As New List(Of Tuple(Of Integer, OcrElement, OcrElement))()
            For Each host In hosts
                Dim tips = ArrowTips(ink, host, tags)
                If tips.Count = 0 Then Continue For
                Dim best As OcrElement = Nothing
                Dim bestD = Integer.MaxValue
                For Each tag In tags
                    If tag Is host Then Continue For
                    ' 球标和尺寸在同一行。落在下一行的数字是另一条尺寸，比如 Rz 下面的 45。
                    If Not SharesCross(host.BoundingBox, tag.BoundingBox) Then Continue For
                    If Not Beside(host.BoundingBox, tag.BoundingBox) Then Continue For
                    Dim dist = DistToTips(tag.BoundingBox, tips)
                    Dim g = Math.Max(MinorSide(host.BoundingBox), MinorSide(tag.BoundingBox))
                    ' 尺寸数字在两头箭头中间。球标贴在某一个箭头尖上，而且要靠近这条尺寸线。
                    If dist > g * 2.2 OrElse Not OnArrowLine(tag.BoundingBox, tips, g) Then Continue For
                    If Not AtTip(tag.BoundingBox, tips, g) Then Continue For
                    If dist < bestD Then
                        bestD = dist
                        best = tag
                    End If
                Next
                If best IsNot Nothing Then ranked.Add(Tuple.Create(bestD, best, host))
            Next

            Dim usedTag As New HashSet(Of OcrElement)()
            Dim usedHost As New HashSet(Of OcrElement)()
            Dim remove As New HashSet(Of OcrElement)()
            Dim add As New List(Of OcrElement)()
            For Each pair In ranked.OrderBy(Function(item) item.Item1)
                Dim tag = pair.Item2
                Dim host = pair.Item3
                If usedTag.Contains(tag) OrElse usedHost.Contains(host) Then Continue For
                usedTag.Add(tag)
                usedHost.Add(host)
                Dim members As New List(Of OcrElement) From {host, tag}
                remove.Add(host)
                remove.Add(tag)
                Dim block = DrawingDimensionGrouper.JoinOnLeader(members)
                add.Add(block)
                System.Diagnostics.Debug.WriteLine("[箭头] " & block.Text)
            Next
            If remove.Count = 0 Then Return
            cells.RemoveAll(Function(el) remove.Contains(el))
            cells.AddRange(add)
        End Sub

        ''' <summary>
        ''' 球标数字：1-2 位普通数字，或单个带圈数字（①..⑳、❶..❿、⓪..⑳、㉑..㉟、㊱..㊿）。
        ''' 带圈数字保持原文：解析脚本的 readBalloon 认 ① 不认裸 1，改写会丢球标序号。
        ''' </summary>
        Private Shared ReadOnly BalloonPiece As New Regex("^(?:\d{1,2}|[\u2460-\u2473\u2776-\u277F\u24EA-\u24FF\u3251-\u325F\u32B1-\u32BF])$", RegexOptions.Compiled)

        Private Class Tip
            Public Property X As Integer
            Public Property Y As Integer
        End Class

        ''' <summary>顺着字框外的细尺寸线走到头。球标盖住箭头时，线在球标外侧停住。</summary>
        Private Shared Function ArrowTips(ink As Boolean(,), host As OcrElement, tags As List(Of OcrElement)) As List(Of Tip)
            Dim found As New List(Of Tip)()
            Dim box = host.BoundingBox
            Dim g = MinorSide(box)
            ' 单个数字的字框常常比字宽更高。只有明显拉长的才是竖排尺寸。
            Dim horizontal = (box.Bottom - box.Top) < (box.Right - box.Left) * 2
            If horizontal Then
                Dim reach = Math.Min(g, 24)
                Dim y0 = Math.Max(1, box.Bottom - CInt(reach * 0.5))
                Dim y1 = Math.Min(ink.GetLength(1) - 2, box.Bottom + CInt(reach * 0.7))
                Dim bestY = Baseline(ink, y0, y1, box.Bottom, box.Left, box.Right, g, True, 0, -1)
                If bestY < 0 Then Return found
                Dim leftTip = WalkArrow(ink, bestY, box.Left, -1, g, True, host, tags)
                Dim rightTip = WalkArrow(ink, bestY, box.Right, 1, g, True, host, tags)
                If leftTip IsNot Nothing Then found.Add(leftTip)
                If rightTip IsNot Nothing Then found.Add(rightTip)
            Else
                Dim x0 = Math.Max(1, box.Left - CInt(g * 0.95))
                Dim x1 = Math.Min(ink.GetLength(0) - 2, box.Right + CInt(g * 0.95))
                Dim bestX = Baseline(ink, x0, x1, (box.Left + box.Right) \ 2, box.Top, box.Bottom, g, False, box.Left + 2, box.Right - 2)
                If bestX < 0 Then Return found
                Dim upTip = WalkArrow(ink, bestX, box.Top, -1, g, False, host, tags)
                Dim downTip = WalkArrow(ink, bestX, box.Bottom, 1, g, False, host, tags)
                If upTip IsNot Nothing Then found.Add(upTip)
                If downTip IsNot Nothing Then found.Add(downTip)
            End If
            Return found
        End Function

        Private Shared Function Baseline(ink As Boolean(,), fromLine As Integer, toLine As Integer, prefer As Integer, origin0 As Integer, origin1 As Integer, g As Integer, horizontal As Boolean, skip0 As Integer, skip1 As Integer) As Integer
            Dim maxScore = 0
            For line = fromLine To toLine
                If line > skip0 AndAlso line < skip1 Then Continue For
                Dim score = ThinRun(ink, line, origin0, -1, horizontal, g * 5) + ThinRun(ink, line, origin1, 1, horizontal, g * 5)
                If score > maxScore Then maxScore = score
            Next
            If maxScore < Math.Max(6, g \ 3) Then Return -1
            Dim best = -1
            Dim bestDist = Integer.MaxValue
            For line = fromLine To toLine
                If line > skip0 AndAlso line < skip1 Then Continue For
                Dim score = ThinRun(ink, line, origin0, -1, horizontal, g * 5) + ThinRun(ink, line, origin1, 1, horizontal, g * 5)
                If score < maxScore * 0.65 Then Continue For
                Dim dist = Math.Abs(line - prefer)
                If dist < bestDist Then
                    bestDist = dist
                    best = line
                End If
            Next
            Return best
        End Function

        Private Shared Function WalkArrow(ink As Boolean(,), line As Integer, origin As Integer, dir As Integer, g As Integer, horizontal As Boolean, host As OcrElement, tags As List(Of OcrElement)) As Tip
            Dim cap = Math.Max(28, Math.Min(g, 32) * 4)
            Dim lastThin = origin
            For i = 1 To cap
                Dim at = origin + dir * i
                Dim px = If(horizontal, at, line)
                Dim py = If(horizontal, line, at)
                If ThinShaft(ink, px, py, horizontal) Then
                    lastThin = at
                    Continue For
                End If
                If Resumes(ink, line, at, dir, horizontal) Then Continue For
                ' 球标和尺寸贴在一起时，细线一出字框就进了圆，前面攒不下细杆。
                Return TipAtEnd(ink, line, lastThin, dir, g, horizontal, host, tags)
            Next
            Return TipAtEnd(ink, line, lastThin, dir, g, horizontal, host, tags)
        End Function

        ''' <summary>细线到头。箭头露在外面就认箭头；箭头埋在球标里，就看线是不是在这个球标外侧停住。</summary>
        Private Shared Function TipAtEnd(ink As Boolean(,), line As Integer, at As Integer, dir As Integer, g As Integer, horizontal As Boolean, host As OcrElement, tags As List(Of OcrElement)) As Tip
            Dim dirX = If(horizontal, dir, 0)
            Dim dirY = If(horizontal, 0, dir)
            Dim tag = TagOnEnd(tags, host, line, at, dir, g, horizontal)
            Dim tipAt = at
            If tag IsNot Nothing Then tipAt = OuterAt(tag, dir, horizontal)
            Dim tx = If(horizontal, tipAt, line)
            Dim ty = If(horizontal, line, tipAt)
            If ArrowNear(ink, tx, ty, dirX, dirY) Then Return New Tip With {.X = tx, .Y = ty}
            If tag IsNot Nothing AndAlso EndsThere(ink, line, tipAt, dir, horizontal, g) Then Return New Tip With {.X = tx, .Y = ty}
            Return Nothing
        End Function

        Private Shared Function TagOnEnd(tags As List(Of OcrElement), host As OcrElement, line As Integer, at As Integer, dir As Integer, g As Integer, horizontal As Boolean) As OcrElement
            Dim best As OcrElement = Nothing
            Dim bestD = Integer.MaxValue
            For Each tag In tags
                If tag Is host OrElse tag.BoundingBox Is Nothing Then Continue For
                If Not SharesCross(host.BoundingBox, tag.BoundingBox) Then Continue For
                Dim box = tag.BoundingBox
                Dim along0 = If(horizontal, box.Left, box.Top)
                Dim along1 = If(horizontal, box.Right, box.Bottom)
                Dim cross0 = If(horizontal, box.Top, box.Left)
                Dim cross1 = If(horizontal, box.Bottom, box.Right)
                If line < cross0 - 6 OrElse line > cross1 + 6 Then Continue For
                Dim nearEdge = If(dir > 0, along0, along1)
                Dim delta = (nearEdge - at) * dir
                Dim inside = at >= along0 - 4 AndAlso at <= along1 + 4
                If Not inside AndAlso (delta < -Math.Max(8, g) OrElse delta > 18) Then Continue For
                Dim apart = If(inside, 0, Math.Abs(delta))
                If apart < bestD Then
                    bestD = apart
                    best = tag
                End If
            Next
            Return best
        End Function

        Private Shared Function OuterAt(tag As OcrElement, dir As Integer, horizontal As Boolean) As Integer
            Dim box = tag.BoundingBox
            If horizontal Then Return If(dir > 0, box.Right, box.Left)
            Return If(dir > 0, box.Bottom, box.Top)
        End Function

        Private Shared Function EndsThere(ink As Boolean(,), line As Integer, tipAt As Integer, dir As Integer, horizontal As Boolean, g As Integer) As Boolean
            Dim window = Math.Max(8, g \ 2)
            Dim hits = 0
            For i = 8 To 8 + window
                If OnLine(ink, line, tipAt + dir * i, horizontal) Then hits += 1
            Next
            Return hits * 3 < window
        End Function

        Private Shared Function Resumes(ink As Boolean(,), line As Integer, at As Integer, dir As Integer, horizontal As Boolean) As Boolean
            For k = 2 To 10
                Dim nxt = at + dir * k
                Dim px = If(horizontal, nxt, line)
                Dim py = If(horizontal, line, nxt)
                If ThinShaft(ink, px, py, horizontal) Then Return True
            Next
            Return False
        End Function

        Private Shared Function ThinRun(ink As Boolean(,), line As Integer, origin As Integer, dir As Integer, horizontal As Boolean, cap As Integer) As Integer
            Dim count = 0
            Dim gap = 0
            For i = 1 To cap
                Dim at = origin + dir * i
                Dim px = If(horizontal, at, line)
                Dim py = If(horizontal, line, at)
                If ThinShaft(ink, px, py, horizontal) Then
                    count += 1
                    gap = 0
                Else
                    gap += 1
                    If gap > 3 Then Exit For
                End If
            Next
            Return count
        End Function

        Private Shared Function ThinShaft(ink As Boolean(,), x As Integer, y As Integer, horizontal As Boolean) As Boolean
            If horizontal Then
                If Not (InkAt(ink, x, y) OrElse InkAt(ink, x, y - 1) OrElse InkAt(ink, x, y + 1)) Then Return False
                If InkAt(ink, x, y - 4) AndAlso InkAt(ink, x, y + 4) Then Return False
                Return True
            End If
            If Not (InkAt(ink, x, y) OrElse InkAt(ink, x - 1, y) OrElse InkAt(ink, x + 1, y)) Then Return False
            If InkAt(ink, x - 4, y) AndAlso InkAt(ink, x + 4, y) Then Return False
            Return True
        End Function

        Private Shared Function ArrowNear(ink As Boolean(,), x As Integer, y As Integer, dirX As Integer, dirY As Integer) As Boolean
            For along = 0 To 6
                For side = -2 To 2
                    Dim px = x - dirX * along + dirY * side
                    Dim py = y - dirY * along + dirX * side
                    If IsArrowAt(ink, px, py, dirX, dirY) Then Return True
                Next
            Next
            Return False
        End Function

        Private Shared Function OnArrowLine(box As OcrBoundingBox, tips As List(Of Tip), g As Integer) As Boolean
            Dim cx = (box.Left + box.Right) \ 2
            Dim cy = (box.Top + box.Bottom) \ 2
            Dim along = CInt(g * 2.2)
            Dim aside = CInt(g * 0.9)
            For Each tip In tips
                If Math.Abs(cx - tip.X) <= along AndAlso Math.Abs(cy - tip.Y) <= aside Then Return True
            Next
            Return False
        End Function

        ''' <summary>球标只靠近其中一端箭头。写在两头中间、到两端差不多远的是尺寸本身。</summary>
        Private Shared Function AtTip(box As OcrBoundingBox, tips As List(Of Tip), g As Integer) As Boolean
            If tips Is Nothing OrElse tips.Count = 0 Then Return False
            Dim cx = (box.Left + box.Right) \ 2
            Dim cy = (box.Top + box.Bottom) \ 2
            Dim nearest = Integer.MaxValue
            Dim second = Integer.MaxValue
            For Each tip In tips
                Dim dx = cx - tip.X
                Dim dy = cy - tip.Y
                Dim dist = CInt(Math.Sqrt(dx * dx + dy * dy))
                If dist < nearest Then
                    second = nearest
                    nearest = dist
                ElseIf dist < second Then
                    second = dist
                End If
            Next
            If nearest > g * 1.8 Then Return False
            If tips.Count >= 2 AndAlso second < nearest * 2.2 AndAlso nearest > g * 0.45 Then Return False
            Return True
        End Function

        Private Shared Function DistToTips(box As OcrBoundingBox, tips As List(Of Tip)) As Integer
            Dim cx = (box.Left + box.Right) \ 2
            Dim cy = (box.Top + box.Bottom) \ 2
            Dim best = Integer.MaxValue
            For Each tip In tips
                Dim dx = cx - tip.X
                Dim dy = cy - tip.Y
                Dim dist = CInt(Math.Sqrt(dx * dx + dy * dy))
                If dist < best Then best = dist
            Next
            Return best
        End Function

        ''' <summary>横排尺寸的球标和字框在上下方向重叠，竖排则在左右方向重叠。</summary>
        Private Shared Function SharesCross(host As OcrBoundingBox, tag As OcrBoundingBox) As Boolean
            If host Is Nothing OrElse tag Is Nothing Then Return False
            Dim horizontal = (host.Bottom - host.Top) < (host.Right - host.Left) * 2
            If horizontal Then
                Dim overlap = Math.Min(host.Bottom, tag.Bottom) - Math.Max(host.Top, tag.Top)
                Dim need = Math.Max(8, Math.Min(host.Bottom - host.Top, tag.Bottom - tag.Top) \ 3)
                Return overlap >= need
            End If
            Dim overlapX = Math.Min(host.Right, tag.Right) - Math.Max(host.Left, tag.Left)
            Dim needX = Math.Max(8, Math.Min(host.Right - host.Left, tag.Right - tag.Left) \ 3)
            Return overlapX >= needX
        End Function

        Private Shared Function IsArrowTag(el As OcrElement) As Boolean
            If el Is Nothing OrElse el.BoundingBox Is Nothing OrElse el.Type = OcrElementType.DatumCircle Then Return False
            If el.Type = OcrElementType.Balloon OrElse IsBalloonLike(el) Then Return True
            Dim text = Compact(el.Text)
            If Not BalloonPiece.IsMatch(text) Then Return False
            Dim w = Math.Max(1, el.BoundingBox.Right - el.BoundingBox.Left)
            Dim h = Math.Max(1, el.BoundingBox.Bottom - el.BoundingBox.Top)
            ' 球标接近圆形。又宽又扁的 45 是尺寸，不是箭头上的球标。
            Dim longSide = Math.Max(w, h)
            Dim shortSide = Math.Min(w, h)
            Return longSide < shortSide * 1.7
        End Function

        Private Shared Sub GroupBySharedLeader(ink As Boolean(,), cells As List(Of OcrElement))
            If ink Is Nothing Then Return
            Dim balloons = cells.Where(Function(el) IsBalloonLike(el)).ToList()
            Dim hosts = cells.Where(Function(el) IsCalloutHost(el) AndAlso Not IsBalloonLike(el)).ToList()
            If balloons.Count = 0 OrElse hosts.Count = 0 Then Return

            Dim ranked As New List(Of Tuple(Of Integer, OcrElement, OcrElement))()
            For Each balloon In balloons
                ' 先沿引线走：引线那头接的是哪个尺寸就绑哪个（与图幅、比例、横竖排无关）
                Dim walked = LeaderWalkTarget(ink, balloon, hosts, balloons)
                If walked IsNot Nothing Then
                    ranked.Add(Tuple.Create(0, balloon, walked))
                    Continue For
                End If
                Dim plain As OcrElement = Nothing
                Dim plainD = Integer.MaxValue
                Dim deviation As OcrElement = Nothing
                Dim deviationD = Integer.MaxValue
                For Each host In hosts
                    ' 球标贴在这条尺寸尾端。隔着别的字、或落到下一行，就是另一条引线。
                    If Not SameBand(host.BoundingBox, balloon.BoundingBox) Then Continue For
                    If Not Beside(host.BoundingBox, balloon.BoundingBox) Then Continue For
                    If Not ShareLeader(ink, balloon, host) Then Continue For
                    Dim dist = CenterDist(balloon.BoundingBox, host.BoundingBox)
                    If IsBareDeviation(Compact(host.Text)) Then
                        If dist < deviationD Then
                            deviationD = dist
                            deviation = host
                        End If
                    ElseIf dist < plainD Then
                        plainD = dist
                        plain = host
                    End If
                Next
                Dim chosen = If(plain IsNot Nothing, plain, deviation)
                Dim chosenD = If(plain IsNot Nothing, plainD, deviationD)
                If chosen IsNot Nothing Then ranked.Add(Tuple.Create(chosenD, balloon, chosen))
            Next

            Dim usedBalloon As New HashSet(Of OcrElement)()
            Dim usedHost As New HashSet(Of OcrElement)()
            Dim remove As New HashSet(Of OcrElement)()
            Dim add As New List(Of OcrElement)()
            For Each pair In ranked.OrderBy(Function(item) item.Item1)
                Dim balloon = pair.Item2
                Dim host = pair.Item3
                If usedBalloon.Contains(balloon) OrElse usedHost.Contains(host) Then Continue For
                usedBalloon.Add(balloon)
                usedHost.Add(host)
                ' 只收这一对，再加上叠在这条尺寸上的偏差。不能顺着长线把旁边的尺寸一起带走。
                Dim members As New List(Of OcrElement) From {host, balloon}
                For Each other In hosts
                    If other Is host OrElse usedHost.Contains(other) Then Continue For
                    If Not StackedOn(host, other) Then Continue For
                    members.Add(other)
                    usedHost.Add(other)
                Next
                For Each el In members
                    remove.Add(el)
                Next
                Dim block = DrawingDimensionGrouper.JoinOnLeader(members)
                add.Add(block)
                System.Diagnostics.Debug.WriteLine("[共用引线] " & block.Text)
            Next
            If remove.Count = 0 Then Return
            cells.RemoveAll(Function(el) remove.Contains(el))
            cells.AddRange(add)
        End Sub

        ''' <summary>
        ''' 沿球标自己的引线墨迹走，看引线那头接的是哪个尺寸：直接走到尺寸文字框，
        ''' 或者走到尺寸线某一端的箭头尖（球标常放在引线末端、离文字老远）。
        ''' 为什么需要它：老路子靠「同带 + 间隔阈值」猜，图纸一缩放就时灵时不灵——
        ''' 用户实测同一张图里有的气泡并进了尺寸框、有的没有。引线是实打实的墨迹，
        ''' 连到哪就绑到哪，与图幅、比例、文字横竖排都无关。
        ''' </summary>
        Private Shared Function LeaderWalkTarget(ink As Boolean(,), balloon As OcrElement, hosts As List(Of OcrElement), tags As List(Of OcrElement)) As OcrElement
            If ink Is Nothing OrElse balloon Is Nothing OrElse balloon.BoundingBox Is Nothing OrElse hosts.Count = 0 Then Return Nothing
            ' 只对验过圆环的球标走引线：框就是圈，探针从圈外出发，不会顺着圈沿打转。
            ' 没验到圈的（文字框紧包数字）走老路子兜底。
            If balloon.Type <> OcrElementType.Balloon Then Return Nothing
            Dim g = Math.Max(12, MinorSide(balloon.BoundingBox))
            Dim path = FollowLeader(ink, balloon.BoundingBox, g)
            If path.Count = 0 Then Return Nothing
            Dim endAt = path(path.Count - 1)
            Dim best As OcrElement = Nothing
            Dim bestD = Integer.MaxValue
            For Each host In hosts
                If host Is balloon OrElse host.BoundingBox Is Nothing Then Continue For
                Dim d = PointBoxDist(endAt, host.BoundingBox)
                Dim tips = ArrowTips(ink, host, tags)
                For Each tip In tips
                    Dim dx = tip.X - endAt.X
                    Dim dy = tip.Y - endAt.Y
                    d = Math.Min(d, CInt(Math.Sqrt(dx * dx + dy * dy)))
                Next
                If d < bestD Then
                    bestD = d
                    best = host
                End If
            Next
            Dim reach = g * 1.6
            If best IsNot Nothing AndAlso bestD <= reach Then Return best
            Return Nothing
        End Function

        Private Shared Function PointBoxDist(p As Point, box As OcrBoundingBox) As Integer
            Dim dx = 0
            If p.X < box.Left Then dx = box.Left - p.X
            If p.X > box.Right Then dx = p.X - box.Right
            Dim dy = 0
            If p.Y < box.Top Then dy = box.Top - p.Y
            If p.Y > box.Bottom Then dy = p.Y - box.Bottom
            Return CInt(Math.Sqrt(dx * dx + dy * dy))
        End Function

        ''' <summary>沿方向 d 从框边向外第 k 个像素点；frac 取 1~3，是垂直轴上的 1/4、1/2、3/4。</summary>
        Private Shared Function ProbePoint(box As OcrBoundingBox, d As Point, frac As Integer, k As Integer) As Point
            If d.X <> 0 Then
                Dim x = If(d.X > 0, box.Right + k, box.Left - k)
                Return New Point(x, box.Top + CInt((box.Bottom - box.Top) * frac / 4))
            End If
            Dim y = If(d.Y > 0, box.Bottom + k, box.Top - k)
            Return New Point(box.Left + CInt((box.Right - box.Left) * frac / 4), y)
        End Function

        ''' <summary>这个点上有一条沿 d 方向的细线。</summary>
        Private Shared Function ThinAlong(ink As Boolean(,), x As Integer, y As Integer, d As Point) As Boolean
            Return ThinShaft(ink, x, y, d.X <> 0)
        End Function

        ''' <summary>
        ''' 从框出发沿引线细墨迹走，返回走过的点（走不动就停）。允许折 1~3 道弯，
        ''' 引线常先横后竖；容忍 2 像素断墨。只在球标圈外打转的当没走到。
        ''' </summary>
        Private Shared Function FollowLeader(ink As Boolean(,), box As OcrBoundingBox, g As Integer) As List(Of Point)
            Dim path As New List(Of Point)()
            If ink Is Nothing OrElse box Is Nothing Then Return path
            Dim probeMax = Math.Max(4, CInt(g * 1.2))

            ' 出口：从框边向外探（上/下/左/右各三条射线），哪边接着细墨迹
            Dim startDir As Point = Nothing
            Dim startAt As Point = Nothing
            Dim startK = 1
            Dim startHits = 0
            Dim dirs As Point() = {New Point(1, 0), New Point(-1, 0), New Point(0, 1), New Point(0, -1)}
            For Each d In dirs
                For frac = 1 To 3
                    Dim kBest = 0
                    Dim hits = 0
                    Dim gap = 0
                    For k = 1 To probeMax
                        Dim p = ProbePoint(box, d, frac, k)
                        If ThinAlong(ink, p.X, p.Y, d) Then
                            hits += 1
                            kBest = k
                            gap = 0
                        Else
                            gap += 1
                            If gap > 2 Then Exit For
                        End If
                    Next
                    If hits * 2 >= probeMax AndAlso hits > startHits Then
                        startHits = hits
                        startDir = d
                        startK = Math.Max(1, kBest)
                        startAt = ProbePoint(box, d, frac, startK)
                    End If
                Next
            Next
            If startDir Is Nothing Then Return path

            Dim cur = startDir
            Dim pos = startAt
            path.Add(pos)
            Dim total = startK
            Dim bends = 0
            While total < g * 20 AndAlso bends <= 3
                ' 直走
                Dim moved = 0
                Dim gap2 = 0
                For k = 1 To CInt(g * 2)
                    Dim nx = pos.X + cur.X * k
                    Dim ny = pos.Y + cur.Y * k
                    If ThinAlong(ink, nx, ny, cur) Then
                        moved = k
                        gap2 = 0
                    Else
                        gap2 += 1
                        If gap2 > 2 Then Exit For
                    End If
                Next
                If moved > 0 Then
                    pos = New Point(pos.X + cur.X * moved, pos.Y + cur.Y * moved)
                    path.Add(pos)
                    total += moved
                    Continue While
                End If
                ' 直走不动了，试着转弯（引线常先横后竖）
                Dim turned = False
                Dim sides As Point() = {New Point(cur.Y, cur.X), New Point(-cur.Y, -cur.X)}
                For Each nd In sides
                    Dim probe = 0
                    Dim pgap = 0
                    For k = 1 To probeMax
                        Dim nx = pos.X + nd.X * k
                        Dim ny = pos.Y + nd.Y * k
                        If ThinAlong(ink, nx, ny, nd) Then
                            probe = k
                            pgap = 0
                        Else
                            pgap += 1
                            If pgap > 2 Then Exit For
                        End If
                    Next
                    If probe >= Math.Max(3, g \ 3) Then
                        pos = New Point(pos.X + nd.X * probe, pos.Y + nd.Y * probe)
                        path.Add(pos)
                        total += probe
                        cur = nd
                        bends += 1
                        turned = True
                        Exit For
                    End If
                Next
                If Not turned Then Exit While
            End While

            ' 走回框里 = 绕着球标圈打转，当没走到
            For i = 1 To path.Count - 1
                Dim p = path(i)
                If p.X > box.Left - 2 AndAlso p.X < box.Right + 2 AndAlso p.Y > box.Top - 2 AndAlso p.Y < box.Bottom + 2 Then
                    path.Clear()
                    Exit For
                End If
            Next
            Return path
        End Function

        ''' <summary>
        ''' 两个框共用一条引线。竖排尺寸的字框又高又窄，距离按短边算，
        ''' 避免把同一条长尺寸线远端的球标也收进来。框叠在一起时，看旁边那条尺寸线。
        ''' </summary>
        Private Shared Function ShareLeader(ink As Boolean(,), a As OcrElement, b As OcrElement) As Boolean
            If a Is Nothing OrElse b Is Nothing OrElse a.BoundingBox Is Nothing OrElse b.BoundingBox Is Nothing Then Return False
            Dim g = GlyphOf(a.BoundingBox, b.BoundingBox)
            If CenterDist(a.BoundingBox, b.BoundingBox) > g * 4.5 Then Return False
            If InkBridge(ink, a.BoundingBox, b.BoundingBox, g, True) Then Return True
            If InkBridge(ink, a.BoundingBox, b.BoundingBox, g, False) Then Return True
            Return SideBySide(a.BoundingBox, b.BoundingBox, g) AndAlso SideLine(ink, a.BoundingBox, b.BoundingBox, g)
        End Function

        Private Shared Function GlyphOf(a As OcrBoundingBox, b As OcrBoundingBox) As Integer
            Return Math.Max(MinorSide(a), MinorSide(b))
        End Function

        Private Shared Function MinorSide(box As OcrBoundingBox) As Integer
            Return Math.Max(8, Math.Min(box.Right - box.Left, box.Bottom - box.Top))
        End Function

        ''' <summary>两个框几乎挨着。Rz 和下一行的 4 中间隔着一整格，不能算同一条引线。</summary>
        Private Shared Function Beside(host As OcrBoundingBox, balloon As OcrBoundingBox) As Boolean
            Dim balloonSide = MinorSide(balloon)
            Dim limit = Math.Max(36, CInt(balloonSide * 2.4))
            Return BoxGap(host, balloon) <= limit
        End Function

        ''' <summary>球标要落在这条尺寸自己的字带里，不能落到下一行。</summary>
        Private Shared Function SameBand(host As OcrBoundingBox, balloon As OcrBoundingBox) As Boolean
            Dim g = Math.Max(MinorSide(host), MinorSide(balloon))
            Dim horizontal = (host.Right - host.Left) >= (host.Bottom - host.Top)
            Dim cx = (balloon.Left + balloon.Right) \ 2
            Dim cy = (balloon.Top + balloon.Bottom) \ 2
            If horizontal Then
                Dim pad = CInt(g * 0.75)
                Return cy >= host.Top - pad AndAlso cy <= host.Bottom + pad
            End If
            Dim xPad = CInt(g * 1.8)
            Dim yPad = CInt(g * 0.8)
            Return cx >= host.Left - xPad AndAlso cx <= host.Right + xPad AndAlso cy >= host.Top - yPad AndAlso cy <= host.Bottom + yPad
        End Function

        Private Shared Function SideBySide(a As OcrBoundingBox, b As OcrBoundingBox, g As Integer) As Boolean
            Dim dx = Math.Abs(((a.Left + a.Right) \ 2) - ((b.Left + b.Right) \ 2))
            Dim dy = Math.Abs(((a.Top + a.Bottom) \ 2) - ((b.Top + b.Bottom) \ 2))
            If dx <= g * 1.7 AndAlso dy <= g * 4.5 AndAlso dy >= g * 0.2 Then Return True
            If dy <= g * 1.7 AndAlso dx <= g * 4.5 AndAlso dx >= g * 0.2 Then Return True
            Return False
        End Function

        ''' <summary>竖排看文字旁边的竖线，横排看文字上下的横线。线要在字框外面。</summary>
        Private Shared Function SideLine(ink As Boolean(,), a As OcrBoundingBox, b As OcrBoundingBox, g As Integer) As Boolean
            Dim dx = Math.Abs(((a.Left + a.Right) \ 2) - ((b.Left + b.Right) \ 2))
            Dim dy = Math.Abs(((a.Top + a.Bottom) \ 2) - ((b.Top + b.Bottom) \ 2))
            Dim vertical = dx <= dy
            Dim span0 = If(vertical, Math.Min((a.Top + a.Bottom) \ 2, (b.Top + b.Bottom) \ 2), Math.Min((a.Left + a.Right) \ 2, (b.Left + b.Right) \ 2))
            Dim span1 = If(vertical, Math.Max((a.Top + a.Bottom) \ 2, (b.Top + b.Bottom) \ 2), Math.Max((a.Left + a.Right) \ 2, (b.Left + b.Right) \ 2))
            If span1 - span0 < 4 Then Return False
            Dim fromLine = If(vertical, Math.Min(a.Left, b.Left), Math.Min(a.Top, b.Top)) - CInt(g * 2.2)
            Dim toLine = If(vertical, Math.Max(a.Right, b.Right), Math.Max(a.Bottom, b.Bottom)) + CInt(g * 2.2)
            For line = fromLine To toLine
                If InsideBox(line, a, vertical) OrElse InsideBox(line, b, vertical) Then Continue For
                Dim hits = 0
                Dim samples = 0
                For pos = span0 To span1 Step 2
                    samples += 1
                    If OnLine(ink, line, pos, Not vertical) Then hits += 1
                Next
                If samples < 3 OrElse hits < samples * 0.55 Then Continue For
                If ThinSpan(ink, line, span0, span1, Not vertical) Then Return True
            Next
            Return False
        End Function

        Private Shared Function InsideBox(line As Integer, box As OcrBoundingBox, vertical As Boolean) As Boolean
            If vertical Then Return line > box.Left + 2 AndAlso line < box.Right - 2
            Return line > box.Top + 2 AndAlso line < box.Bottom - 2
        End Function

        Private Shared Function IsBalloonLike(el As OcrElement) As Boolean
            If el Is Nothing OrElse el.BoundingBox Is Nothing Then Return False
            If el.Type = OcrElementType.DatumCircle Then Return False
            If el.Type = OcrElementType.Balloon Then Return True
            Dim text = Compact(el.Text)
            If Not BalloonPiece.IsMatch(text) Then Return False
            Dim w = Math.Max(1, el.BoundingBox.Right - el.BoundingBox.Left)
            Dim h = Math.Max(1, el.BoundingBox.Bottom - el.BoundingBox.Top)
            Return h >= w * 1.1 AndAlso h <= w * 2.05
        End Function

        Private Shared Function InkBridge(ink As Boolean(,), a As OcrBoundingBox, b As OcrBoundingBox, h As Integer, horizontal As Boolean) As Boolean
            Dim first = a
            Dim second = b
            If horizontal Then
                If first.Left > second.Left Then
                    first = b
                    second = a
                End If
            ElseIf first.Top > second.Top Then
                first = b
                second = a
            End If
            Dim gapStart = If(horizontal, first.Right, first.Bottom)
            Dim gapEnd = If(horizontal, second.Left, second.Top)
            If gapEnd - gapStart < 3 OrElse gapEnd - gapStart > h * 3 Then Return False

            Dim fromAt = Math.Max(0, If(horizontal, Math.Min(first.Top, second.Top), Math.Min(first.Left, second.Left)) - h \ 2)
            Dim toAt = If(horizontal, Math.Max(first.Bottom, second.Bottom), Math.Max(first.Right, second.Right)) + h \ 2
            Dim limit = If(horizontal, ink.GetLength(1), ink.GetLength(0)) - 2
            If toAt > limit Then toAt = limit
            For at = fromAt To toAt
                If EdgeDist(at, first, horizontal) > h * 1.25 OrElse EdgeDist(at, second, horizontal) > h * 1.25 Then Continue For
                Dim hits = 0
                Dim samples = 0
                For pos = gapStart To gapEnd
                    samples += 1
                    If OnLine(ink, at, pos, horizontal) Then hits += 1
                Next
                If samples < 4 OrElse hits < samples * 0.5 Then Continue For
                Dim mid0 = gapStart + samples \ 5
                Dim mid1 = gapEnd - samples \ 5
                If mid1 <= mid0 Then
                    mid0 = gapStart
                    mid1 = gapEnd
                End If
                If ThinSpan(ink, at, mid0, mid1, horizontal) Then Return True
            Next
            Return False
        End Function

        Private Shared Function EdgeDist(at As Integer, box As OcrBoundingBox, horizontal As Boolean) As Integer
            Dim near = If(horizontal, box.Top, box.Left)
            Dim far = If(horizontal, box.Bottom, box.Right)
            If at >= near AndAlso at <= far Then Return 0
            If at < near Then Return near - at
            Return at - far
        End Function

        Private Shared Function ThinSpan(ink As Boolean(,), line As Integer, startAt As Integer, finishAt As Integer, horizontal As Boolean) As Boolean
            Dim thin = 0
            Dim seen = 0
            For at = startAt To finishAt Step 2
                seen += 1
                Dim hit = OnLine(ink, line, at, horizontal)
                Dim above = If(horizontal, InkAt(ink, at, line - 3), InkAt(ink, line - 3, at))
                Dim below = If(horizontal, InkAt(ink, at, line + 3), InkAt(ink, line + 3, at))
                If hit AndAlso Not above AndAlso Not below Then thin += 1
            Next
            If seen < 2 Then Return True
            Return thin / CDbl(seen) >= 0.4
        End Function

        Private Shared Function StackedOn(host As OcrElement, other As OcrElement) As Boolean
            If other Is Nothing OrElse other.BoundingBox Is Nothing OrElse host.BoundingBox Is Nothing Then Return False
            If Not IsBareDeviation(Compact(other.Text)) Then Return False
            Dim h = MinorSide(host.BoundingBox)
            Dim overlap = Math.Min(host.BoundingBox.Right, other.BoundingBox.Right) - Math.Max(host.BoundingBox.Left, other.BoundingBox.Left)
            If overlap < Math.Max(4, (other.BoundingBox.Right - other.BoundingBox.Left) \ 3) Then Return False
            Dim dy = Math.Abs(((host.BoundingBox.Top + host.BoundingBox.Bottom) \ 2) - ((other.BoundingBox.Top + other.BoundingBox.Bottom) \ 2))
            Return dy <= CInt(h * 1.45)
        End Function

        Private Shared Function IsCalloutHost(el As OcrElement) As Boolean
            If el Is Nothing OrElse el.BoundingBox Is Nothing Then Return False
            If el.Type = OcrElementType.Balloon OrElse el.Type = OcrElementType.DatumCircle Then Return False
            Dim text = Compact(el.Text)
            If text = "" OrElse text.Length > 40 OrElse text.StartsWith("FCF:") Then Return False
            Return Regex.IsMatch(text, "\d")
        End Function

        Private Shared Function IsBareDeviation(text As String) As Boolean
            Return Regex.IsMatch(text, "^(?:[±+-]\d+(?:\.\d+)?|0)$")
        End Function

        Private Shared Function IsSeparateCallout(text As String) As Boolean
            If text = "" OrElse IsBareDeviation(text) Then Return False
            If Regex.IsMatch(text, "^(?:Ra|Rz|Rq|Rt|Rh|Sm)\d", RegexOptions.IgnoreCase) Then Return True
            If Regex.IsMatch(text, "^(?:[A-Z]{1,2}\d{1,2}|[a-z]{1,2}\d{1,2})$") Then Return False
            If Regex.IsMatch(text, "^(?:SØ|SR|Ø|ø|Φ|φ|⌀|∅|□|R|C|M)$") Then Return False
            Return Regex.IsMatch(text, "\d")
        End Function

        Private Shared Function CenterDist(a As OcrBoundingBox, b As OcrBoundingBox) As Integer
            Dim dx = ((a.Left + a.Right) \ 2) - ((b.Left + b.Right) \ 2)
            Dim dy = ((a.Top + a.Bottom) \ 2) - ((b.Top + b.Bottom) \ 2)
            Return CInt(Math.Sqrt(dx * dx + dy * dy))
        End Function

        Private Shared Function Bind(ink As Boolean(,), seed As OcrElement, cells As List(Of OcrElement), horizontal As Boolean) As List(Of OcrElement)
            Dim landing = If(horizontal, FindHorizontal(ink, seed), FindVertical(ink, seed))
            If landing Is Nothing Then Return Nothing
            Dim members = Collect(ink, seed, cells, landing)
            If members Is Nothing Then Return Nothing
            Dim h = Math.Max(8, seed.BoundingBox.Bottom - seed.BoundingBox.Top)
            Dim balloon = BalloonOnLanding(seed, cells, landing, h)
            If balloon IsNot Nothing AndAlso Not members.Contains(balloon) Then members.Add(balloon)
            If members.Count < 2 Then Return Nothing
            Return members
        End Function

        Private Shared Function FindHorizontal(ink As Boolean(,), seed As OcrElement) As Landing
            Dim box = seed.BoundingBox
            Dim h = Math.Max(8, box.Bottom - box.Top)
            Dim height = ink.GetLength(1)
            Dim y0 = Math.Max(1, box.Bottom - CInt(h * 0.3))
            Dim y1 = Math.Min(height - 2, box.Bottom + CInt(h * 1.05))
            Dim bestY = -1
            Dim bestDist = Integer.MaxValue
            For y = y0 To y1
                Dim leftLen = Probe(ink, y, box.Left, -1, 0, CInt(h * 12))
                Dim rightLen = Probe(ink, y, box.Right, 1, 0, CInt(h * 12))
                If leftLen < h * 0.35 AndAlso rightLen < h * 0.35 Then Continue For
                Dim dist = Math.Abs(y - box.Bottom)
                If dist < bestDist Then
                    bestDist = dist
                    bestY = y
                End If
            Next
            If bestY < 0 Then Return Nothing
            Return FinishLanding(ink, seed, bestY, True)
        End Function

        Private Shared Function FindVertical(ink As Boolean(,), seed As OcrElement) As Landing
            Dim box = seed.BoundingBox
            Dim h = Math.Max(8, box.Bottom - box.Top)
            Dim width = ink.GetLength(0)
            Dim x0 = Math.Max(1, box.Left - CInt(h * 1.15))
            Dim x1 = Math.Min(width - 2, box.Right + CInt(h * 1.15))
            Dim bestX = -1
            Dim bestDist = Integer.MaxValue
            For x = x0 To x1
                If x >= box.Left AndAlso x <= box.Right Then Continue For
                Dim upLen = Probe(ink, x, box.Top, 0, -1, CInt(h * 12))
                Dim downLen = Probe(ink, x, box.Bottom, 0, 1, CInt(h * 12))
                If upLen < h * 0.35 AndAlso downLen < h * 0.35 Then Continue For
                Dim dist = Math.Min(Math.Abs(x - box.Left), Math.Abs(x - box.Right))
                If dist < bestDist Then
                    bestDist = dist
                    bestX = x
                End If
            Next
            If bestX < 0 Then Return Nothing
            Return FinishLanding(ink, seed, bestX, False)
        End Function

        Private Shared Function FinishLanding(ink As Boolean(,), seed As OcrElement, line As Integer, horizontal As Boolean) As Landing
            Dim box = seed.BoundingBox
            Dim h = Math.Max(8, box.Bottom - box.Top)
            Dim origin = If(horizontal, (box.Left + box.Right) \ 2, (box.Top + box.Bottom) \ 2)
            Dim backX = If(horizontal, -1, 0)
            Dim backY = If(horizontal, 0, -1)
            Dim foreX = If(horizontal, 1, 0)
            Dim foreY = If(horizontal, 0, 1)
            Dim cap = CInt(h * 12)
            Dim startAt = Extend(ink, line, origin, backX, backY, cap)
            Dim finishAt = Extend(ink, line, origin, foreX, foreY, cap)
            If finishAt - startAt < Math.Max(10, CInt(h * 0.8)) Then Return Nothing
            Dim sticks As Boolean
            If horizontal Then
                sticks = startAt < box.Left - h * 0.3 OrElse finishAt > box.Right + h * 0.3
            Else
                sticks = startAt < box.Top - h * 0.3 OrElse finishAt > box.Bottom + h * 0.3
            End If
            If Not sticks Then Return Nothing
            If Not ThinLanding(ink, seed, line, startAt, finishAt, horizontal) Then Return Nothing

            ' 折线的箭头在斜段末端，直段本身可以没有箭头。两端都没有箭头就不是引线。
            Dim arrowBack = IsArrow(ink, startAt, line, backX, backY, horizontal) OrElse BendToArrow(ink, startAt, line, backX, backY, h, horizontal)
            Dim arrowFore = IsArrow(ink, finishAt, line, foreX, foreY, horizontal) OrElse BendToArrow(ink, finishAt, line, foreX, foreY, h, horizontal)
            If Not arrowBack AndAlso Not arrowFore Then Return Nothing

            ' 箭头在线段中途时，把落点收到箭头之间，避免把下一条尺寸一起吃进来。
            If arrowBack Then startAt = ArrowStop(ink, line, origin, backX, backY, cap, horizontal)
            If arrowFore Then finishAt = ArrowStop(ink, line, origin, foreX, foreY, cap, horizontal)
            If finishAt - startAt < Math.Max(8, CInt(h * 0.5)) Then Return Nothing

            Return New Landing With {.Horizontal = horizontal, .Line = line, .Start = startAt, .Finish = finishAt}
        End Function

        Private Shared Function ArrowStop(ink As Boolean(,), line As Integer, origin As Integer, dirX As Integer, dirY As Integer, cap As Integer, horizontal As Boolean) As Integer
            Dim along = If(horizontal, dirX, dirY)
            Dim at = origin
            For i = 1 To cap
                Dim nextAt = at + along
                If Not OnLine(ink, line, nextAt, horizontal) Then Exit For
                at = nextAt
                If i > 4 AndAlso IsArrow(ink, at, line, dirX, dirY, horizontal) Then Return at
            Next
            Return at
        End Function

        Private Shared Function Collect(ink As Boolean(,), seed As OcrElement, cells As List(Of OcrElement), landing As Landing) As List(Of OcrElement)
            Dim h = Math.Max(8, seed.BoundingBox.Bottom - seed.BoundingBox.Top)
            Dim members As New List(Of OcrElement) From {seed}
            ' 整块兜底：一条引线最多一条完整标注。CanJoin 只拿候选和种子逐个比，
            ' 挡不住「裸偏差两边各贴着一条尺寸」——偏差居中时两边都像叠在它身上。
            Dim hasCallout = IsSeparateCallout(Compact(seed.Text))
            For Each other In cells
                If other Is seed OrElse Not CanJoin(seed, other, h) Then Continue For
                If Not OnLanding(seed, other, landing, h) Then Continue For
                If Separated(ink, seed, other, landing) Then Continue For
                If IsSeparateCallout(Compact(other.Text)) Then
                    If hasCallout Then Continue For
                    hasCallout = True
                End If
                members.Add(other)
            Next
            If members.Count > 6 Then Return Nothing
            Return members
        End Function

        Private Shared Function OnLanding(seed As OcrElement, other As OcrElement, landing As Landing, h As Integer) As Boolean
            Dim box = other.BoundingBox
            Dim cx = (box.Left + box.Right) \ 2
            Dim cy = (box.Top + box.Bottom) \ 2
            Dim pad = CInt(h * 0.45)
            If landing.Horizontal Then
                If cx < landing.Start - pad OrElse cx > landing.Finish + pad Then Return False
                If cy < landing.Line - CInt(h * 1.85) OrElse cy > landing.Line + CInt(h * 0.75) Then Return False
            Else
                If cy < landing.Start - pad OrElse cy > landing.Finish + pad Then Return False
                If Math.Abs(cx - landing.Line) > CInt(h * 2.1) Then Return False
            End If
            Dim sx = (seed.BoundingBox.Left + seed.BoundingBox.Right) \ 2
            Dim sy = (seed.BoundingBox.Top + seed.BoundingBox.Bottom) \ 2
            Dim along = If(landing.Horizontal, Math.Abs(cx - sx), Math.Abs(cy - sy))
            If along > h * 3 Then Return False
            Return StackedOrBeside(seed, other, h)
        End Function

        ''' <summary>球标在引线尾端或尺寸旁边。同一段上只收离这条尺寸最近的一个。</summary>
        Private Shared Function BalloonOnLanding(seed As OcrElement, cells As List(Of OcrElement), landing As Landing, h As Integer) As OcrElement
            Dim best As OcrElement = Nothing
            Dim bestApart As Integer = Integer.MaxValue
            Dim pad = CInt(h * 2.4)
            For Each other In cells
                If other Is Nothing OrElse other Is seed OrElse other.Type <> OcrElementType.Balloon OrElse other.BoundingBox Is Nothing Then Continue For
                Dim box = other.BoundingBox
                Dim cx = (box.Left + box.Right) \ 2
                Dim cy = (box.Top + box.Bottom) \ 2
                Dim onSpan As Boolean
                If landing.Horizontal Then
                    onSpan = cx >= landing.Start - pad AndAlso cx <= landing.Finish + pad AndAlso cy >= landing.Line - CInt(h * 1.6) AndAlso cy <= landing.Line + CInt(h * 0.45)
                Else
                    onSpan = cy >= landing.Start - pad AndAlso cy <= landing.Finish + pad AndAlso Math.Abs(cx - landing.Line) <= CInt(h * 3.2)
                End If
                If Not onSpan Then Continue For
                Dim apart = BoxGap(seed.BoundingBox, box)
                If apart < bestApart Then
                    bestApart = apart
                    best = other
                End If
            Next
            If best Is Nothing Then Return Nothing
            Dim balloonSide = MinorSide(best.BoundingBox)
            If bestApart > Math.Max(36, CInt(balloonSide * 2.4)) Then Return Nothing
            Return best
        End Function

        ''' <summary>偏差叠在名义尺寸上；前缀和公差等级贴在名义尺寸旁边。同一行再往外的尺寸不收。</summary>
        Private Shared Function StackedOrBeside(seed As OcrElement, other As OcrElement, h As Integer) As Boolean
            Dim a = seed.BoundingBox
            Dim b = other.BoundingBox
            Dim yOverlap = Math.Min(a.Bottom, b.Bottom) - Math.Max(a.Top, b.Top)
            Dim xOverlap = Math.Min(a.Right, b.Right) - Math.Max(a.Left, b.Left)
            If yOverlap > h * 0.35 Then
                Dim gap = 0
                If b.Left > a.Right Then gap = b.Left - a.Right
                If a.Left > b.Right Then gap = a.Left - b.Right
                Return gap <= h * 0.8
            End If
            Dim smaller = Math.Max(1, Math.Min(a.Right - a.Left, b.Right - b.Left))
            If xOverlap >= smaller * 0.3 Then Return True
            Dim acx = (a.Left + a.Right) \ 2
            Return acx >= b.Left - h * 0.4 AndAlso acx <= b.Right + h * 0.4
        End Function

        Private Shared Function Separated(ink As Boolean(,), seed As OcrElement, other As OcrElement, landing As Landing) As Boolean
            Dim a = seed.BoundingBox
            Dim b = other.BoundingBox
            If landing.Horizontal Then
                Dim yFrom = Math.Min(b.Bottom, landing.Line) + 2
                Dim yTo = Math.Max(b.Top, landing.Line) - 2
                If yTo - yFrom < 3 Then Return False
                For y = yFrom To yTo
                    If y = landing.Line Then Continue For
                    If RunSticksPast(ink, y, Math.Min(a.Left, b.Left), Math.Max(a.Right, b.Right), True) Then Return True
                Next
            Else
                Dim xFrom = Math.Min(b.Right, landing.Line) + 2
                Dim xTo = Math.Max(b.Left, landing.Line) - 2
                If xTo - xFrom < 3 Then Return False
                For x = xFrom To xTo
                    If x = landing.Line Then Continue For
                    If RunSticksPast(ink, x, Math.Min(a.Top, b.Top), Math.Max(a.Bottom, b.Bottom), False) Then Return True
                Next
            End If
            Return False
        End Function

        Private Shared Function CanJoin(seed As OcrElement, other As OcrElement, h As Integer) As Boolean
            If other Is Nothing OrElse other.BoundingBox Is Nothing Then Return False
            If other.Type = OcrElementType.Balloon OrElse other.Type = OcrElementType.DatumCircle Then Return False
            Dim text = Compact(other.Text)
            If text = "" OrElse text.Length > 24 Then Return False
            If text.StartsWith("FCF:") Then Return False
            If Regex.IsMatch(text, "^[A-Z]$") Then Return False
            ' 一条引线上最多收一条完整标注（名义值），外加叠在它身上的公差栈。
            ' 旧写法只在「种子和候选都是完整标注」时拦截：种子是裸偏差（+0.2）或丢符号的
            ' 单个数字时守卫形同虚设，同一条尺寸线上挨着的 49-0、352-0、Rz25 会被一起
            ' 扫进一块——识别盒子就是这样脱离引线规则的。这里按「块里只能有一条完整标注」挡。
            If IsSeparateCallout(text) Then
                If IsSeparateCallout(Compact(seed.Text)) Then Return False
                If Not StackedOnSeed(seed, other) Then Return False
            End If
            If Not Regex.IsMatch(text, "\d") AndAlso Not Regex.IsMatch(text, "^(?:SØ|SR|Ø|ø|Φ|φ|⌀|∅|□|R|C|M)$") Then Return False
            ' 尾端单独的一两位数字多半是球标，不是这条尺寸的偏差。
            If Regex.IsMatch(text, "^\d{1,2}$") Then
                Dim box = other.BoundingBox
                Dim side = Math.Max(box.Right - box.Left, box.Bottom - box.Top)
                Dim shortSide = Math.Max(1, Math.Min(box.Right - box.Left, box.Bottom - box.Top))
                Dim square = side <= shortSide * 1.8
                Dim apart = BoxGap(seed.BoundingBox, box)
                If square AndAlso apart > h * 0.35 Then Return False
            End If
            Return True
        End Function

        Private Shared Function IsSeed(el As OcrElement) As Boolean
            If el Is Nothing OrElse el.BoundingBox Is Nothing Then Return False
            If el.Type = OcrElementType.Balloon OrElse el.Type = OcrElementType.DatumCircle Then Return False
            Dim text = Compact(el.Text)
            If text = "" OrElse text.Length > 24 OrElse text.StartsWith("FCF:") Then Return False
            Return Regex.IsMatch(text, "\d")
        End Function

        ''' <summary>
        ''' 候选框叠在种子正上/正下方（公差栈），而不是并排的另一条尺寸。
        ''' 上下偏差写在名义值两边，中心线基本对齐；并排的另一条尺寸中心明显错开。
        ''' </summary>
        Private Shared Function StackedOnSeed(seed As OcrElement, other As OcrElement) As Boolean
            If seed Is Nothing OrElse other Is Nothing OrElse seed.BoundingBox Is Nothing OrElse other.BoundingBox Is Nothing Then Return False
            If Not IsBareDeviation(Compact(seed.Text)) AndAlso Not IsBareDeviation(Compact(other.Text)) Then Return False
            Dim a = seed.BoundingBox
            Dim b = other.BoundingBox
            Dim yOverlap = Math.Min(a.Bottom, b.Bottom) - Math.Max(a.Top, b.Top)
            Dim h = Math.Max(8, a.Bottom - a.Top)
            If yOverlap >= h * 0.3 Then Return False
            Dim dx = Math.Abs(((a.Left + a.Right) \ 2) - ((b.Left + b.Right) \ 2))
            Dim span = Math.Max(a.Right - a.Left, b.Right - b.Left)
            Return dx <= Math.Max(8, CInt(span * 0.75))
        End Function

        Private Shared Function Probe(ink As Boolean(,), line As Integer, origin As Integer, dirX As Integer, dirY As Integer, cap As Integer) As Integer
            Dim along = If(dirX <> 0, dirX, dirY)
            Dim count = 0
            Dim gap = 0
            Dim at = origin
            For i = 1 To cap
                at += along
                If OnLine(ink, line, at, dirX <> 0) Then
                    count += 1
                    gap = 0
                Else
                    gap += 1
                    If gap > 3 Then Exit For
                End If
            Next
            Return count
        End Function

        Private Shared Function Extend(ink As Boolean(,), line As Integer, origin As Integer, dirX As Integer, dirY As Integer, cap As Integer) As Integer
            Dim along = If(dirX <> 0, dirX, dirY)
            Dim at = origin
            Dim gap = 0
            For i = 1 To cap
                Dim nextAt = at + along
                If OnLine(ink, line, nextAt, dirX <> 0) Then
                    at = nextAt
                    gap = 0
                Else
                    gap += 1
                    If gap > 4 Then Exit For
                    at = nextAt
                End If
            Next
            Return at - along * gap
        End Function

        Private Shared Function OnLine(ink As Boolean(,), line As Integer, at As Integer, horizontal As Boolean) As Boolean
            If horizontal Then
                Return InkAt(ink, at, line) OrElse InkAt(ink, at, line - 1) OrElse InkAt(ink, at, line + 1)
            End If
            Return InkAt(ink, line, at) OrElse InkAt(ink, line - 1, at) OrElse InkAt(ink, line + 1, at)
        End Function

        Private Shared Function CountAlong(ink As Boolean(,), line As Integer, fromAt As Integer, toAt As Integer, horizontal As Boolean) As Integer
            Dim hits = 0
            For at = fromAt To toAt
                If OnLine(ink, line, at, horizontal) Then hits += 1
            Next
            Return hits
        End Function

        Private Shared Function ThinLanding(ink As Boolean(,), seed As OcrElement, line As Integer, startAt As Integer, finishAt As Integer, horizontal As Boolean) As Boolean
            Dim box = seed.BoundingBox
            Dim thin = 0
            Dim seen = 0
            For at = startAt To finishAt Step 2
                If horizontal Then
                    If at >= box.Left AndAlso at <= box.Right Then Continue For
                ElseIf at >= box.Top AndAlso at <= box.Bottom Then
                    Continue For
                End If
                seen += 1
                Dim hit = OnLine(ink, line, at, horizontal)
                Dim above = If(horizontal, InkAt(ink, at, line - 3), InkAt(ink, line - 3, at))
                Dim below = If(horizontal, InkAt(ink, at, line + 3), InkAt(ink, line + 3, at))
                If hit AndAlso Not above AndAlso Not below Then thin += 1
            Next
            If seen < 4 Then Return True
            Return thin / CDbl(seen) >= 0.35
        End Function

        Private Shared Function RunSticksPast(ink As Boolean(,), line As Integer, fromAt As Integer, toAt As Integer, horizontal As Boolean) As Boolean
            Dim left = Probe(ink, line, fromAt, If(horizontal, -1, 0), If(horizontal, 0, -1), 80)
            Dim right = Probe(ink, line, toAt, If(horizontal, 1, 0), If(horizontal, 0, 1), 80)
            Return left > 6 OrElse right > 6
        End Function

        Private Shared Function IsArrow(ink As Boolean(,), tipAlong As Integer, line As Integer, dirX As Integer, dirY As Integer, horizontal As Boolean) As Boolean
            Dim x = If(horizontal, tipAlong, line)
            Dim y = If(horizontal, line, tipAlong)
            Return IsArrowAt(ink, x, y, dirX, dirY)
        End Function

        Private Shared Function IsArrowAt(ink As Boolean(,), x As Integer, y As Integer, dirX As Integer, dirY As Integer) As Boolean
            Dim sx = Math.Sign(dirX)
            Dim sy = Math.Sign(dirY)
            If sx = 0 AndAlso sy = 0 Then Return False
            Dim earlyMax = 0
            Dim lateMax = 0
            Dim tipW = 0
            Dim sideA = 0
            Dim sideB = 0
            Dim bulgeA = 0
            Dim bulgeB = 0
            For i = 0 To 12
                Dim px = x - sx * i
                Dim py = y - sy * i
                Dim w = WidthAround(ink, px, py, sx, sy, sideA, sideB)
                If i = 0 Then tipW = w
                If i <= 7 Then
                    If w > earlyMax Then
                        earlyMax = w
                        bulgeA = sideA
                        bulgeB = sideB
                    End If
                ElseIf w > lateMax Then
                    lateMax = w
                End If
            Next
            ' 箭头两侧都有墨。框角只在线的一侧鼓起。
            If bulgeA = 0 OrElse bulgeB = 0 Then Return False
            Return earlyMax >= 4 AndAlso earlyMax <= 14 AndAlso tipW <= earlyMax - 2 AndAlso earlyMax >= lateMax + 2 AndAlso lateMax <= 3
        End Function

        Private Shared Function BendToArrow(ink As Boolean(,), tipAlong As Integer, line As Integer, dirX As Integer, dirY As Integer, h As Integer, horizontal As Boolean) As Boolean
            Dim slopes = New Integer() {1, -1, 2, -2}
            For Each slope In slopes
                Dim cx = If(horizontal, tipAlong, line)
                Dim cy = If(horizontal, line, tipAlong)
                Dim stepX = If(horizontal, dirX, slope)
                Dim stepY = If(horizontal, slope, dirY)
                Dim steps = 0
                For n = 1 To Math.Max(8, CInt(h * 7))
                    Dim nx = cx + stepX
                    Dim ny = cy + stepY
                    If InkAt(ink, nx, ny) Then
                        cx = nx
                        cy = ny
                        steps += 1
                    ElseIf InkAt(ink, nx + Math.Sign(stepX), ny) OrElse InkAt(ink, nx, ny + Math.Sign(stepY)) Then
                        cx = nx
                        cy = ny
                        steps += 1
                    Else
                        Exit For
                    End If
                Next
                If steps < Math.Max(6, CInt(h * 0.6)) Then Continue For
                If IsArrowAt(ink, cx, cy, Math.Sign(stepX), Math.Sign(stepY)) Then Return True
            Next
            Return False
        End Function

        Private Shared Function WidthAround(ink As Boolean(,), x As Integer, y As Integer, dirX As Integer, dirY As Integer, ByRef sideA As Integer, ByRef sideB As Integer) As Integer
            Dim px = If(dirY = 0, 0, -Math.Sign(dirY))
            Dim py = If(dirX = 0, 0, Math.Sign(dirX))
            If dirX <> 0 AndAlso dirY <> 0 Then
                px = -dirY
                py = dirX
            ElseIf dirX = 0 Then
                px = 1
                py = 0
            Else
                px = 0
                py = 1
            End If
            If InkAt(ink, x + px * 12, y + py * 12) AndAlso InkAt(ink, x - px * 12, y - py * 12) Then Return 1
            Dim w = 0
            sideA = 0
            sideB = 0
            For t = 1 To 8
                If InkAt(ink, x + px * t, y + py * t) Then
                    w += 1
                    sideA += 1
                End If
                If InkAt(ink, x - px * t, y - py * t) Then
                    w += 1
                    sideB += 1
                End If
            Next
            If InkAt(ink, x, y) Then w += 1
            Return w
        End Function

        Private Shared Function InkAt(ink As Boolean(,), x As Integer, y As Integer) As Boolean
            If ink Is Nothing OrElse x < 0 OrElse y < 0 OrElse x >= ink.GetLength(0) OrElse y >= ink.GetLength(1) Then Return False
            Return ink(x, y)
        End Function

        Private Shared Function BoxGap(a As OcrBoundingBox, b As OcrBoundingBox) As Integer
            Dim dx = 0
            If b.Left > a.Right Then dx = b.Left - a.Right
            If a.Left > b.Right Then dx = a.Left - b.Right
            Dim dy = 0
            If b.Top > a.Bottom Then dy = b.Top - a.Bottom
            If a.Top > b.Bottom Then dy = a.Top - b.Bottom
            Return Math.Max(dx, dy)
        End Function

        Private Shared Function Compact(text As String) As String
            Return Regex.Replace(If(text, ""), "\s+", "")
        End Function

        Private Shared Function LoadInk(source As Bitmap) As Boolean(,)
            Dim w = source.Width
            Dim h = source.Height
            If w < 8 OrElse h < 8 OrElse w > 8000 OrElse h > 8000 Then Return Nothing
            Dim ink(w - 1, h - 1) As Boolean
            Using copy As New Bitmap(w, h, PixelFormat.Format24bppRgb)
                Using g As Graphics = Graphics.FromImage(copy)
                    g.DrawImage(source, 0, 0, w, h)
                End Using
                Dim data = copy.LockBits(New Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)
                Try
                    Dim stride = data.Stride
                    Dim count = stride * h
                    Dim bytes(count - 1) As Byte
                    Marshal.Copy(data.Scan0, bytes, 0, count)
                    For y = 0 To h - 1
                        Dim row = y * stride
                        For x = 0 To w - 1
                            Dim i = row + x * 3
                            Dim lum = (CInt(bytes(i + 2)) * 299 + CInt(bytes(i + 1)) * 587 + CInt(bytes(i)) * 114) \ 1000
                            ink(x, y) = lum < 170
                        Next
                    Next
                Finally
                    copy.UnlockBits(data)
                End Try
            End Using
            Return ink
        End Function

    End Class

End Namespace
