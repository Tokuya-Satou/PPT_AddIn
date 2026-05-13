# リボンUI実装アイデア比較

作成日: 2026-04-17

---

## 前提: 既存コードの現状

| 要素 | 現状 |
|---|---|
| `DragDropRibbon.xml` | 右クリックメニューのみ（タブなし） |
| `DragDropRibbon.cs` | `ribbon.Invalidate()` 済み、コールバック構造あり |
| `PenPaletteWindow.xaml.cs` | 色がハードコード（Black/Red/Blue/Yellow固定） |
| アドイン ON/OFF | 手段なし（スライドショー開始で常に起動） |

---

## 案A: 多機能構成（ideas_20260417.md ベース）

### グループ1: 基本制御 (General Control)

| コントロール | 詳細 |
|---|---|
| `[Toggle]` アドイン有効化 | スライドショー時のオーバーレイ全般を制御 |
| `[Button]` 図形名の一括変更 | 選択した図形の名前の先頭に `Drag_` を付加/削除 |
| `[Button]` このプレゼンでは無効 | プレゼンテーション個別のOFF設定 |

### グループ2: ペンパレット設定 (Pen Settings)

| コントロール | 詳細 |
|---|---|
| `[Gallery]` ペンの色 | Office 標準に近いカラーギャラリー。「最近使った色」も保持 |
| `[Menu]` 線の太さ | 1pt / 3pt / 5pt などの選択 |
| `[Toggle]` スライド別記録 | 描画内容をスライドごとに保持するかどうかの切り替え |

### グループ3: パレット表示設定 (Overlay Settings)

| コントロール | 詳細 |
|---|---|
| `[Menu]` パレット位置 | デフォルトの表示位置（左下、右下など）の指定 |
| `[Toggle]` ミニマムモード | 不要なボタン（進む・戻るなど）を非表示にする |

### 案Aの技術方針

- **Ribbon XML**: Visual Studio のデザイナーではなく XML で記述（動的な色変更に必須）
- **設定の永続化**: CustomXMLParts を使用してプレゼンテーションファイル内に保存
- **プレゼンごとの個別設定**: CustomXMLParts 経由でファイルに紐付けて保存
- **ペンプロファイル**: 「赤・太め」「青・マーカー」などのセットをリボンで登録し、パレット側は「Pen 1 / 2 / 3」ボタンのみのシンプル構成

---

## 案B: シンプル版（最小コスト・最大効果）

**原則: 「今の痛み」だけ解決する**

### グループ1: 制御 (Control)

| コントロール | 実装方針 | 理由 |
|---|---|---|
| `[Toggle]` アドイン有効化 | `Properties.Settings` に bool 保存（レジストリ） | シンプル、アプリ全体で共有 |
| `[Button]` Drag_ 付与/解除 | 既存の `OnToggleDrag` をそのまま流用 | コンテキストメニューとロジック共有 |
| `[Button]` このプレゼンでは無効 | `CustomDocumentProperties` に bool 保存 | CustomXMLParts より大幅にシンプル |

**自動判定（案Bの追加アイデア）:**
`Drag_` 図形が1つもないプレゼンでは、トグルに関わらずオーバーレイを起動しない。
`SlideShowBegin` のタイミングで `Shapes` を走査するだけ、追加コストほぼゼロ。

`CustomDocumentProperties` の実装例:
```csharp
// 書き込み
pres.CustomDocumentProperties.Add("PPTDragAddIn_Disabled", false,
    Office.MsoDocProperties.msoPropertyTypeBoolean, true);
// 読み込み
var prop = pres.CustomDocumentProperties["PPTDragAddIn_Disabled"];
```

### グループ2: ペン (Pen)

カラーギャラリーの代わりに **色ボタン6個 + 太さメニュー**:

```
[● 黒][● 赤][● 青][● 緑][■ 蛍光][消しゴム]
[▼ 太さ: 細/中/太]
```

### 案B の Ribbon XML 全体構成

```xml
<tab id="tabPPTDrag" label="PPTドラッグ">

  <group id="grpControl" label="制御">
    <toggleButton id="tglEnabled"
                  label="アドイン有効"
                  getPressed="GetAddinEnabled"
                  onAction="OnToggleAddin" />
    <button id="btnToggleDrag"
            label="Drag_ 設定"
            getLabel="GetDragLabel"
            onAction="OnToggleDrag" />
    <button id="btnDisableForThis"
            label="このプレゼンでは無効"
            onAction="OnDisableForPresentation" />
  </group>

  <group id="grpPen" label="ペン">
    <button id="btnBlack"  label="黒"   onAction="OnPenColor" tag="Black" />
    <button id="btnRed"    label="赤"   onAction="OnPenColor" tag="Red" />
    <button id="btnBlue"   label="青"   onAction="OnPenColor" tag="Blue" />
    <button id="btnGreen"  label="緑"   onAction="OnPenColor" tag="Green" />
    <button id="btnYellow" label="蛍光" onAction="OnPenMarker" />
    <button id="btnEraser" label="消去" onAction="OnEraser" />
    <menu   id="mnuThick"  label="太さ">
      <button label="細" onAction="OnThick" tag="2" />
      <button label="中" onAction="OnThick" tag="5" />
      <button label="太" onAction="OnThick" tag="10" />
    </menu>
  </group>

</tab>
```

---

## 案A vs 案B 比較

| 項目 | 案A（多機能） | 案B（シンプル） |
|---|---|---|
| **設定の永続化** | CustomXMLParts（複雑） | CustomDocumentProperties（シンプル） |
| **色選択UI** | カラーギャラリー（画像リソース要） | 色ボタン6個（XMLのみ） |
| **ペンプロファイル** | あり（Pen 1/2/3） | なし（後のフェーズ） |
| **スライド別記録トグル** | あり | なし（既存動作のまま） |
| **パレット位置設定** | あり | なし（既存動作のまま） |
| **ミニマムモード** | あり | なし（既存動作のまま） |
| **実装量** | 大 | 小 |
| **既存コードの再利用** | 少 | 大（OnToggleDrag 等） |
| **リスク** | XML+画像+状態管理が複雑 | 低い |

### 推奨方針

**案Bを先に実装し、動作確認後に案Aの要素を段階的に追加する。**

優先順位:
1. アドイン ON/OFF トグル（最も使い勝手に直結）
2. Drag_ 付与/解除ボタン（コンテキストメニューと共有）
3. ペン色ボタン（パレットのハードコード解消）
4. 太さメニュー
5. このプレゼンでは無効（CustomDocumentProperties）
6. カラーギャラリー・ペンプロファイル（案Aの要素、後のフェーズ）
