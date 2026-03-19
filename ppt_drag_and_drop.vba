' ==========================================
'  PowerPointドラッグ&ドロップ (環境依存排除版)
' ==========================================

' 名前衝突を避けるため、独自の接頭辞を使用しています。
' モジュールの中身をすべて入れ替えてお試しください。

#If VBA7 Then
    ' 64bit/32bit 両対応
    Private Declare PtrSafe Function TOKUYA_GetPos Lib "user32" Alias "GetCursorPos" (lpPoint As TOKUYA_POINT) As Long
    Private Declare PtrSafe Function TOKUYA_KeyState Lib "user32" Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function TOKUYA_GetPos Lib "user32" Alias "GetCursorPos" (lpPoint As TOKUYA_POINT) As Long
    Private Declare Function TOKUYA_KeyState Lib "user32" Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer
#End If

Private Type TOKUYA_POINT
    X As Long
    Y As Long
End Type

' ドラッグ実行
Sub DragObject(shp As Shape)
    ' 1. 【デバッグ】音が鳴らなければ設定ミスです
    Beep
    
    Dim ssv As SlideShowView
    On Error Resume Next
    Set ssv = ActivePresentation.SlideShowWindow.View
    On Error GoTo 0
    If ssv Is Nothing Then Exit Sub

    ' 2. 初期設定
    Dim m As TOKUYA_POINT
    Dim ratioX As Double, ratioY As Double
    Dim px0 As Long, py0 As Long
    Dim offX As Single, offY As Single
    Dim curX As Single, curY As Single
    
    ' 座標変換の起点と比率をPPTから直接取得
    px0 = ssv.PointToScreenPixelsX(0)
    py0 = ssv.PointToScreenPixelsY(0)
    
    ' 100ポイント間のピクセル数から正確な比率を算出
    ratioX = (ssv.PointToScreenPixelsX(100) - px0) / 100
    ratioY = (ssv.PointToScreenPixelsY(100) - py0) / 100
    
    ' マウス位置を取得し、図形を「中心」で掴む設定
    TOKUYA_GetPos m
    offX = shp.Width / 2
    offY = shp.Height / 2
    
    ' 3. メインループ
    ' &H8000 は「今押されている」というフラグです。
    ' &H1 はマウスの左ボタンです。
    
    Dim i As Long
    ' 最大600回(数十秒)ループ。安全のため回数を設定。
    For i = 1 To 2000
        ' マウス現在位置を取得
        TOKUYA_GetPos m
        
        ' スライド上の座標に変換
        curX = (m.X - px0) / ratioX
        curY = (m.Y - py0) / ratioY
        
        ' 図形を移動
        shp.Left = curX - offX
        shp.Top = curY - offY
        
        ' 画面更新 (最前面に一瞬持っていくことで描画を促す)
        shp.ZOrder msoBringToFront
        DoEvents
        
        ' もし指が離されたら終了
        If (TOKUYA_KeyState(&H1) And &H8000) = 0 Then Exit For
    Next i
    
End Sub
