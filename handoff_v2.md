# PowerPoint ドラッグ＆ドロップアドイン v2.0 (WPFオーバーレイ方式) 開発ドキュメント

## 1. プロジェクトの前提と現状 (v1.0: VSTO 最適化版)
PowerPoint VSTO を使用し、スライドショー中に図形をドラッグ＆ドロップできるアドインを実装・最適化した。

* **技術スタック:** C# (.NET Framework 4.7.2), VSTO (Visual Studio Tools for Office)
* **フック:** Windows Low-Level Mouse Hook (`WH_MOUSE_LL`) を使用し、PowerPoint の「クリック」を遮断して「ドラッグ」として扱う。
* **実装済みの最適化:** COM 通信のキャッシュ、`Stopwatch` による 10ms (100fps) 単位の時間制御、スクリーン座標からスライド座標への変換（アスペクト比対応の黒帯オフセット計算含む）、`try-catch` による安定性ガード。

## 2. 次世代の構想 (v2.0: WPF オーバーレイ方式)
ClassPoint 等と同様の「完全に吸い付く」滑らかさを実現するための新設計。PowerPoint の図形を直接動かすのではなく、スライドショーの真上に透明な WPF ウィンドウを重ね、その上で画像を動かす。

* **オーバーレイ:** 透明な WPF ウィンドウをスライドショーウィンドウ (`SlideShowWindow.HWND`) の真上に、サイズを同期させて配置する。
* **ドラッグ開始:** 対象図形の「スナップショット（画像）」を取得し、WPF ウィンドウ上のマウス位置に表示。PowerPoint 本体の図形は一時的に `Visible = msoFalse` にする。
* **ドラッグ中:** WPF 内の画像 (`Image` コントロール) のみを動かす。ネイティブなフレームレートで極めて滑らかに動作する。
* **ドロップ時:** マウスを離した座標からスライド座標を逆算し、PowerPoint 本体の図形をその位置へ移動させ可視化する。WPF 上の画像は消去する。

### ⚠️ 注意が必要なポイント (ハマりどころ)
1.  **マルチモニターと DPI:** モニターごとに DPI（スケーリング）が異なる場合、`GetWindowRect` で取得するピクセル座標とマウスイベントの座標がズレる。
2.  **スライドショーウィンドウの特定:** 発表者ツールを使用している場合、`Application.SlideShowWindows` から「実際に投影されているウィンドウ（全画面）」を正しく選別する必要がある。
3.  **アスペクト比計算:** スライド (4:3) と画面 (16:9) が異なる際の計算式:
    ```csharp
    float scale = Math.Min(winW / slideW, winH / slideH);
    float slidePixelW = slideW * scale;
    float offsetX = (winW - slidePixelW) / 2;
    ```

### 継承すべき重要ファイル
* `ThisAddIn.cs`: 座標計算と PowerPoint オブジェクト操作。
* `MouseHook.cs`: 低レベルマウスフックの定型コード。
* `DragDropRibbon.xml/cs`: ドラッグ対象を識別するための「Drag_」接頭辞の切り替え。

---

## 3. 開発手順と現在地

開発は以下のStepで進行する。現在は **Step 1 と Step 2 が完了（または作業中）** であり、次は **Step 3** の実装が求められている。

* **[完了] Step 1: プロジェクト基盤の整備 (VSTO + WPFの統合)**
    * VSTOプロジェクトの参照設定に `PresentationCore`, `PresentationFramework`, `WindowsBase`, `System.Xaml` を追加済み。
* **[完了] Step 2: 透明オーバーレイウィンドウの基礎実装**
    * 透明なWPFウィンドウ `OverlayWindow.xaml` を作成済み。
    * プロパティ設定: `WindowStyle="None"`, `AllowsTransparency="True"`, `Background="Transparent"`, `Topmost="True"`, `ShowInTaskbar="False"`, `ResizeMode="NoResize"`。
* **[完了] Step 3: スライドショーウィンドウの捕捉と完全同期**
    * `Application.SlideShowBegin` をフック。
    * 投影用ウィンドウ (`ppSlideShowFullScreen`) の HWND と座標を取得し、WPF ウィンドウと同期。
* **[完了] Step 4:** 「掴む」アクションと画像キャプチャ
    * `shape.Export` によるスナップショット取得と WPF 表示を実装。
* **[完了] Step 5:** 滑らかなドラッグ処理とドロップ時の座標逆算
    * 低レベルマウスフックによる 100fps 級の滑らかな移動と、ドロップ時の座標確定を実装。
* **[完了] Step 6:** マルチモニター・DPIスケーリングの補正（最適化）
    * WPF `OnDpiChanged` による動的なスケール補正と、複数モニター跨ぎの座標計算を実装。

---

## 4. ペンパレット機能の追加 (2026-03-29 完了)

スライドショー中に使えるフローティング・ペンパレットを追加した。

### 実装方式: WPF InkCanvas オーバーレイ

PowerPoint COM の `DrawColor` プロパティがいかなる方法でもアクセス不可と確定したため、PowerPoint のペン機能は使わず **`OverlayWindow` 上の `InkCanvas` で自前描画** する方式を採用。

```
[OverlayWindow (透明 WPF)]
  ├── Canvas         ← ドラッグスナップショット用（既存）
  └── InkCanvas      ← ペン描画用（新規追加）
        Background: alpha=0 → click-through（カーソルモード）
        Background: alpha=1 → イベント受取（描画モード）
```

### 重要な知見: レイヤードウィンドウの click-through 挙動

`AllowsTransparency=True` の WPF ウィンドウでは **alpha=0 ピクセルは OS レベルで click-through** になる。`Background="Transparent"` の InkCanvas はマウスイベントを受け取れないため、描画モード時は `Color.FromArgb(1, 0, 0, 0)`（alpha=1）の背景を設定する必要がある。

### 追加ファイル・変更ファイル

| ファイル | 内容 |
|---|---|
| `PenPaletteWindow.xaml/.cs` | フローティングパレット UI。WM_MOUSEACTIVATE=MA_NOACTIVATE でフォーカス奪取を防止 |
| `OverlayWindow.xaml/.cs` | InkCanvas 追加、SetPenMode/SetEraserMode/SetArrowMode/ClearDrawing を実装 |
| `ThisAddIn.cs` | `_isDrawModeActive` フラグ、公開メソッド群、SlideShowNextSlide で描画クリア |

### 動作仕様

- スライドショー開始時にパレットが自動表示（左下）
- 黒・赤・青ペン、黄色（太め）、消しゴム（ストローク単位）、カーソル切替
- スライド切替時に描画を自動クリア
- 描画モード中はドラッグ操作を無効化（競合防止）
- スライドショー終了時にパレット非表示、描画クリア

---

## 5. ペンパレット機能の改良・バグ修正 (2026-04-01 完了)

### 5-1. パレット UI の改良

#### 最小化/再表示トグル
タイトルバー右端に「−」ボタンを追加。クリックでボタン群を折りたたみ（`PanelContent.Visibility = Collapsed`）、「＋」で再展開。`_penPaletteWindow` は毎回 `new` で再作成するため、スライドショー開始時は必ず展開状態になる。

#### スライドナビゲーション（次へ / 戻る）
パレット下部に「◀ 戻る」「次へ ▶」ボタンを追加。`ThisAddIn.GoPrevSlide()` / `GoNextSlide()` が `_activeShowWindow.View.Previous()` / `.Next()` を呼び出す。

### 5-2. スライドごとの描画保持

スライドを移動しても描画が消えないよう、`Dictionary<int, StrokeCollection>` でスライドインデックスごとにストロークを保存・復元する。

```csharp
// ThisAddIn.cs
private Dictionary<int, StrokeCollection> _slideStrokes;
private int _currentSlideIndex;

// SlideShowNextSlide イベントで保存 → 復元
_slideStrokes[_currentSlideIndex] = _overlayWindow.GetStrokes();  // 保存
_overlayWindow.SetStrokes(_slideStrokes[newIndex]);                // 復元
```

`OverlayWindow` に `GetStrokes()` / `SetStrokes()` を追加。スライドショー終了時にディクショナリをクリア。

### 5-3. バグ修正: 描画モード中にパレットを操作できない

**原因:** 描画モード時に InkCanvas の `Background = Color.FromArgb(1,0,0,0)` がスクリーン全体をカバーし、OS のヒットテストでオーバーレイウィンドウが優先されてパレットウィンドウにクリックが届かなかった。

**解決策:** `OverlayWindow` の `WndProc` で `WM_NCHITTEST` を処理し、パレット領域では `HTTRANSPARENT` を返す。

```csharp
// OverlayWindow.xaml.cs
private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
{
    if (msg == WM_NCHITTEST && _isInDrawMode)
    {
        int x = (short)(lParam.ToInt32() & 0xFFFF);
        int y = (short)((lParam.ToInt32() >> 16) & 0xFFFF);
        if (IsScreenPointOverPalette(x, y))
        {
            handled = true;
            return new IntPtr(HTTRANSPARENT); // OS がこのウィンドウを無視して下のウィンドウへ
        }
    }
    return IntPtr.Zero;
}
```

`HTTRANSPARENT` を返すことで OS がオーバーレイを無視し、パレットウィンドウにイベントが届く。カーソル表示もパレット側の矢印に変わる。

### 5-4. バグ修正: ドラッグ&ドロップが動作しない（クロススレッド問題）

**原因:** マウスフックのコールバックスレッドから `Dispatcher.Invoke` を使わずに WPF オブジェクトを直接操作しており、スレッド違反で `catch` に握りつぶされていた。

**解決策:** `ShowSnapshot`・`HideSnapshot` を `Dispatcher.Invoke`、`UpdatePosition` を `Dispatcher.BeginInvoke` でラップ。また COM 操作（`shape.Visible = msoFalse`）は WPF 操作より先に行う。

**ドラッグ動作の前提条件（重要）:**
- 動かしたい図形の名前を PowerPoint 上で `Drag_` で始まるよう設定すること
- パレットがカーソルモード（↖）になっていること（ペンモード中はドラッグ無効）

---

**AIへの指示:**
上記のコンテキストを読み込み、現在の状況を把握してください。全 Step、ペンパレット機能、およびその改良・バグ修正がすべて完了済みです。
