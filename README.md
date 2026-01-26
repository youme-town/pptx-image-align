# pptx-image-align

![GUI screenshot](docs/images/gui.png)
Example of GUI

![Output example](docs/images/output.png)
Example of output.pptx

複数の画像を PowerPoint（`.pptx`）上に **グリッド / フロー（詰め配置）**で並べて出力するツールです。
GUI でプレビューしながら調整するか、YAML 設定ファイルを使って CLI で一括生成できます。

- CLI: [cli.py](cli.py)
- GUI: [gui.py](gui.py)
- 共通ロジック: [core.py](core.py)

---

## 実装されている機能一覧

### 1. スライド設定

- **スライドサイズ**: 幅・高さを cm 単位で指定（デフォルト: 16:9 比率 33.867×19.05 cm）

### 2. グリッド構成

- **グリッドサイズ**: rows × cols で指定
- **配列モード**: `row`（各フォルダが行）または `col`（各フォルダが列）
- **レイアウトモード**:
  - `grid`: 厳密に整列（ガイド線に沿う）
  - `flow`: 詰め配置（コンパクト）

### 3. フロー配置オプション

- **水平方向揃え**: `left` / `center` / `right`
- **垂直方向揃え**: `top` / `center` / `bottom`

### 4. マージン・間隔設定

- **マージン**: 左・上・右・下を個別に cm 単位で指定
- **間隔（gap）**:
  - 水平・垂直方向を個別に設定
  - **cm モード**: 固定値（cm）
  - **scale モード**: 画像サイズに対する比率

### 5. 画像サイズ設定

- **サイズモード**:
  - `fit`: グリッドに合わせて自動計算
  - `fixed`: 固定サイズを指定
- **フィットモード**（fitモード時）:
  - `fit`: 縦横両方が収まる最大サイズ
  - `width`: 幅を基準にフィット
  - `height`: 高さを基準にフィット
- **スケール**: 画像の拡大/縮小率

### 6. クロップ（拡大表示）機能

#### 基本機能

- **複数クロップ領域**: 1つの画像に対して複数のクロップ領域を定義可能
- **座標指定方式**:
  - `px`（ピクセル）: x, y, width, height を直接指定
  - `ratio`（比率）: x_ratio, y_ratio, width_ratio, height_ratio で指定（サイズの異なる画像に対応）

#### 適用セルの制御

- **rows/cols フィルタ**: 適用する行/列を指定（0-indexed）
- **targets**: 明示的に (row, col) ペアで適用セルを指定
- **overrides**: セル単位で独自のクロップ領域を設定（または無効化）

#### 表示設定

- **配置位置**: `right`（右）または `bottom`（下）
- **サイズ指定**:
  - `size`: 絶対サイズ（cm）
  - `scale`: メイン画像に対する比率
- **アライメント**: `auto` / `start` / `center` / `end`
- **オフセット**: 位置の微調整（cm）
- **個別ギャップ**: クロップ領域ごとにメイン画像との間隔を上書き

#### 間隔設定

- **main_crop_gap**: メイン画像とクロップ画像の間隔
- **crop_crop_gap**: クロップ画像同士の間隔
- **crop_bottom_gap**: クロップ領域下部の追加間隔

### 7. 枠線設定

- **クロップ枠線**（元画像上）:
  - 表示/非表示
  - 線幅（pt）
  - 形状: `rectangle` / `rounded`
- **拡大画像枠線**:
  - 表示/非表示
  - 線幅（pt）
  - 形状: `rectangle` / `rounded`
- **色**: クロップ領域ごとに RGB で指定可能

### 8. 入力方式

- **フォルダモード**: フォルダパスを指定し、中の画像を自動読み込み（ファイル名の数字でソート）
- **画像リストモード**: 画像パスを個別に指定（row-major 順）

### 9. GUI 機能

- **リアルタイムプレビュー**: 設定変更を即座に反映
- **ダミー画像比率設定**: プレビュー用のアスペクト比を指定
- **クロップエディタ**: 画像上でドラッグしてクロップ領域を選択
- **セルクリック**: プレビュー上のセルをクリックしてクロップ編集
- **設定の保存/読み込み**: YAML 形式で設定をエクスポート/インポート
- **タブ構成**:
  - 基本・フォルダ
  - レイアウト
  - クロップ設定
  - 装飾（枠線）

### 10. CLI 機能

- **設定ファイルから生成**: `cli.py config.yaml`
- **サンプル設定生成**: `cli.py --init [filename]`
- **ヘルプ表示**: `cli.py --help`

### 11. 対応画像形式

- PNG, JPG, JPEG, GIF, BMP, TIFF, WebP

---

## 必要環境

- Python >= 3.10（開発設定: [.python-version](.python-version)）
- 依存関係は [pyproject.toml](pyproject.toml) で管理
  - `python-pptx`: PowerPoint 生成
  - `pyyaml`: YAML 設定ファイル解析
  - `pillow`: 画像処理（GUI での表示・クロップ）

---

## インストール（uv 推奨）

### 1) uv の導入

未導入なら（各 OS に合わせて）: <https://docs.astral.sh/uv/>

### 2) 依存関係の同期

```sh
uv sync
```

---

## 使い方（CLI）

### サンプル設定ファイルを生成

```sh
uv run python cli.py --init config.yaml
```

`config.yaml` は [`core.generate_sample_config`](core.py) が生成します。

### YAML から PPTX を生成

```sh
uv run python cli.py config.yaml
```

内部的には以下が使われます：

- 設定ロード: [`core.load_config`](core.py)
- PPTX 生成: [`core.create_grid_presentation`](core.py)

---

## 使い方（GUI）

### 起動

```sh
uv run python gui.py
```

設定ファイルを指定して起動：

```sh
uv run python gui.py config.yaml
```

GUI では、フォルダ追加・レイアウト調整・クロップ領域作成（エディタ）・設定保存・PPTX 生成ができます。

---

## 入力フォルダ構造と並び順

`folders:` に列挙したフォルダから画像を読み込み、ファイル名中の数字でソートします（[`core.get_sorted_images`](core.py)）。

例（本リポジトリの同梱例）:

```text
images/
  row1/
  row2/
```

- `grid.arrangement: row` の場合：各フォルダが「行」
- `grid.arrangement: col` の場合：各フォルダが「列」

---

## 設定ファイル（YAML）概要

サンプル: [config.yaml](config.yaml)

### 主要項目

| カテゴリ | 項目 | 説明 |
| ------- | ---- | ---- |
| `slide` | `width`, `height` | スライドサイズ（cm） |
| `grid` | `rows`, `cols` | グリッドサイズ |
| `grid` | `arrangement` | `row` / `col` |
| `grid` | `layout_mode` | `grid` / `flow` |
| `grid` | `flow_align` | `left` / `center` / `right` |
| `grid` | `flow_vertical_align` | `top` / `center` / `bottom` |
| `margin` | `left`, `top`, `right`, `bottom` | マージン（cm） |
| `gap` | `horizontal`, `vertical` | 間隔（cm or `{value, mode}`） |
| `image` | `size_mode` | `fit` / `fixed` |
| `image` | `fit_mode` | `fit` / `width` / `height` |
| `image` | `width`, `height` | 固定サイズ時のサイズ |
| `crop.regions` | - | クロップ領域リスト |
| `crop` | `rows`, `cols` | 適用セル指定（0-index、`null` は全て） |
| `crop` | `targets` | 適用セル明示指定（例: `[{row:0,col:1}]`） |
| `crop` | `overrides` | セル単位のクロップ上書き |
| `crop.display` | `position` | `right` / `bottom` |
| `crop.display` | `size`, `scale` | サイズ指定 |
| `border.crop` | `show`, `width`, `shape` | クロップ枠線設定 |
| `border.zoom` | `show`, `width`, `shape` | 拡大画像枠線設定 |
| `folders` | - | 入力フォルダリスト |
| `images` | - | 画像パス直接指定（row-major） |
| `output` | - | 出力ファイルパス |

---

## クロップ（拡大表示）について

- クロップ領域はピクセル座標（`x, y, width, height`）で定義できます（[`core.CropRegion`](core.py)）。
- 画像サイズがバラバラでも同じ位置を切り出したい場合、`mode: ratio` と `x_ratio/y_ratio/width_ratio/height_ratio` で比率指定できます。
- 適用セルは [`core.should_apply_crop`](core.py) の条件（`crop.targets` → `crop.overrides` → `crop.rows/cols`）で決まります。
- 拡大画像の配置計算は [`core.calculate_item_bounds`](core.py) と関連関数群で行っています。

---

## アーキテクチャ

```text
┌─────────────┐     ┌─────────────┐
│   gui.py    │     │   cli.py    │
│  (tkinter)  │     │  (argparse) │
└──────┬──────┘     └──────┬──────┘
       │                   │
       └─────────┬─────────┘
                 ▼
         ┌─────────────┐
         │   core.py   │
         │ (共通ロジック) │
         └──────┬──────┘
                 │
    ┌────────────┼────────────┐
    ▼            ▼            ▼
┌─────────┐ ┌─────────┐ ┌─────────┐
│ python- │ │  PIL/   │ │  PyYAML │
│  pptx   │ │ Pillow  │ │         │
└─────────┘ └─────────┘ └─────────┘
```

### 処理フロー

```text
設定ファイル (YAML) or GUI入力
    ↓
load_config() / GUI設定
    ↓
build_image_grid() → rows×cols の画像グリッドを構築
    ↓
calculate_grid_metrics() → 列幅・行高さ・クロップサイズを計算
    ↓
For each cell:
    ├─ calculate_item_bounds() → 画像とクロップの領域を計算
    ├─ add_picture() → PowerPoint にメイン画像を追加
    ├─ add_crop_borders_to_image() → クロップ枠線を描画
    ├─ crop_image() → クロップ領域を抽出
    ├─ add_picture() → クロップ画像を追加
    └─ add_border_shape() → 拡大画像の枠線を追加
    ↓
save() → output.pptx を生成
```

---

## 改善案（反映状況）

- [x] サイズの違う画像を並べる（比率違い含む）
  → `fit_mode` + レイアウト計算で吸収
- [x] サイズの違う画像（縦横比は同じ）の一括クロップ
  → `crop.regions[].mode: ratio` を追加
- [x] クロップできる画像を選択できるようにする
  → `crop.targets` を追加
- [x] 画像をフォルダで追加するだけでなく、画像を個別に追加する
  → GUI に「画像リスト」入力モードを追加（`images` も保存可能）
- [x] プレビュー画面クリックで画像を編集できるようにする（クロップなど）
  → セルクリックで CropEditor を開く
- [ ] プレビュー画面で実画像を表示する（完全プレビュー）

---

## 追加機能の提案

以下は今後の開発で検討できる機能です：

### 優先度: 高

1. **完全プレビュー機能**
   - 現在はダミー画像でプレビューしているが、実画像を表示
   - 出力結果をより正確に確認可能

2. **マルチスライド対応**
   - 画像数がグリッドを超える場合、自動的に複数スライドに分割
   - ページ送り設定（rows × cols × pages）

3. **画像の並び替え（GUI）**
   - ドラッグ＆ドロップで画像の順序を変更
   - 上下移動ボタン

### 優先度: 中

1. **テキストラベル/キャプション**
   - 各画像やクロップ領域にラベルを追加
   - フォント、サイズ、位置のカスタマイズ

2. **ドラッグ＆ドロップ入力**
   - GUI にファイル/フォルダをドラッグ＆ドロップで追加

3. **クロッププリセット**
   - よく使うクロップ設定を保存/読み込み
   - 「3分割」「4分割」などのテンプレート

4. **Undo/Redo 機能**
   - GUI での操作履歴を管理
   - Ctrl+Z / Ctrl+Y で元に戻す/やり直し

### 優先度: 低

1. **テンプレート PPTX 対応**
   - 既存の PPTX をテンプレートとして使用
   - マスタースライドの継承

2. **追加エクスポート形式**
   - PDF 出力
   - PNG/JPG 画像出力（スライド単位）

3. **バッチ処理**
   - 複数の設定ファイルを一括処理
   - コマンドラインからのバッチ実行

4. **クロップ領域の連結線**
   - 元画像のクロップ枠と拡大画像を線で結ぶ
   - 視覚的にどこをクロップしたか明確化

5. **画像フィルタ/エフェクト**
   - グレースケール、セピア、明るさ調整など
   - クロップ画像のみにエフェクト適用

---

## ライセンス

MIT License（[LICENSE](LICENSE)）
