# PPT作成スキル定義（PPT Worker Skills）

## 役割
読み込み資料（XMind・Markdown・テキスト等）をもとに、日本語のPowerPointファイル（.pptx）を生成する。

---

## 使用ライブラリ
- `python-pptx`（PowerPointファイル生成）
- `zipfile` + `json`（XMindファイルの解析）

---

## XMind → PPT 変換手順

### Step 1: XMindファイルの読み込み
XMindファイル（.xmind）はZIP形式。以下の方法で内容を取得する：

```python
import zipfile, json

with zipfile.ZipFile("読み込み資料/ファイル名.xmind", "r") as z:
    with z.open("content.json") as f:
        data = json.load(f)
```

### Step 2: マインドマップ構造の抽出
```python
# ルートトピックとサブトピックを再帰的に取得
def extract_topics(node, level=0):
    title = node.get("title", "")
    children = node.get("children", {}).get("attached", [])
    return {"title": title, "level": level, "children": [extract_topics(c, level+1) for c in children]}
```

### Step 3: PPTスライドの生成

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
# スライドサイズ（ワイド 16:9）
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
```

---

## スライド構成ルール

| スライド | 内容 | 対応ノード |
|---------|------|----------|
| 1枚目 | タイトルスライド | ルートトピック |
| 2枚目以降 | 各メインブランチ | Level 1ノード |
| 箇条書き | サブ項目 | Level 2以下ノード |

---

## デザインルール（日本語対応）

- **フォント**：游ゴシック または メイリオ（日本語対応フォント）
- **タイトル文字サイズ**：36pt
- **本文文字サイズ**：24pt
- **配色**：ネイビー（#003366）+ ホワイト + ライトグレー
- **余白**：上下左右 0.5インチ以上

---

## 出力ルール

| フェーズ | 保存先 | ファイル名例 |
|---------|--------|------------|
| 初回生成 | `出力/下書き/` | `下書き_v1.pptx` |
| 修正版 | `出力/下書き/` | `下書き_v2.pptx` |
| 最終承認後 | `出力/最終版/` | `最終版_YYYYMMDD.pptx` |

---

## 作業開始前チェックリスト
- [ ] `読み込み資料/` にXMindファイルが格納されている
- [ ] XMindのルートトピック・階層構造を確認済み
- [ ] スライド枚数の目安を確認済み（1メインブランチ = 1スライドが基本）
- [ ] 日本語フォントが指定されている
