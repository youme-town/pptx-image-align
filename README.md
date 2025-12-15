# pptx-image-align

![GUI screenshot](docs/images/gui.png)
Example of GUI

![Output example](docs/images/output.png)
Example of output.pptx

複数の画像を PowerPoint（`.pptx`）上に **グリッド / フロー（詰め配置）**で並べて出力するツールです。  
GUI でプレビューしながら調整するか、YAML 設定ファイルを使って CLI で一括生成できます。

- CLI: [cli.py](cli.py)
- GUI: [gui.py](gui.py)
- 共通ロジック: [core.py](core.py)（例: [`core.create_grid_presentation`](core.py), [`core.load_config`](core.py)）

---

## 特徴

- グリッド配置 / フロー配置（`grid.layout_mode`）
- フロー時の水平方向揃え（left/center/right）・垂直揃え（top/center/bottom）
- 余白・間隔（cm または scale）
- 画像サイズ: fit / fixed、fit モード（fit/width/height）
- クロップ領域を複数定義し、拡大画像を右または下に配置
- 元画像上のクロップ枠線／拡大画像の枠線（rectangle / rounded）
- YAML 設定の保存/読込（GUI/CLI 両方）

---

## 必要環境

- Python >= 3.10（開発設定: [.python-version](.python-version)）
- 依存関係は [pyproject.toml](pyproject.toml) で管理  
  - `python-pptx`, `pyyaml`, `pillow`（GUI で `PIL` を使用）

---

## インストール（uv 推奨）

### 1) uv の導入
未導入なら（各 OS に合わせて）: https://docs.astral.sh/uv/

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

主要項目：
- `slide.width`, `slide.height`（cm）
- `grid.rows`, `grid.cols`, `grid.arrangement`
- `grid.layout_mode`: `grid` / `flow`
- `grid.flow_align`, `grid.flow_vertical_align`（flow のみ）
- `margin.*`（cm）
- `gap.horizontal`, `gap.vertical`（cm or `{value, mode}`）
- `image.size_mode`: `fit` / `fixed`
- `image.fit_mode`: `fit` / `width` / `height`
- `crop.regions`: 複数クロップ領域（px）
- `crop.rows`, `crop.cols`: 適用セル指定（0-index、`null` は全て）
- `crop.display.position`: `right` / `bottom`
- `border.crop`, `border.zoom`

---

## クロップ（拡大表示）について

- クロップ領域はピクセル座標（`x, y, width, height`）で定義します（[`core.CropRegion`](core.py)）。
- 適用セルは [`core.should_apply_crop`](core.py) の条件（`crop.rows/cols`）で決まります。
- 拡大画像の配置計算は [`core.calculate_item_bounds`](core.py) と関連関数群で行っています。

## ライセンス

MIT License（[LICENSE](LICENSE)）