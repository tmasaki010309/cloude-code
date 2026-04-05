# PPT作成スキル定義（PPT Worker Skills）

## 役割
読み込み資料（XMind・PNG・Markdown・テキスト等）をもとに、日本語のPowerPointファイル（.pptx）を生成する。
マインドマップの内容をそのまま転記するのではなく、**プラスアルファのコンテンツ（統計・事例・Tipsなど）を補完**して資料の価値を高める。

---

## 使用ライブラリ
- `python-pptx`（PowerPointファイル生成）
- `zipfile` + `json`（XMindファイルの解析）
- `lxml`（XML操作・角丸矩形等）

---

## フォント
- **必須フォント：Meiryo UI**（日本語対応・Windows標準）
- フォールバック：メイリオ
- 英語部分も統一してMeiryo UIを使用する

---

## デザインルール（GenSpark風フォーマット）

### カラーパレット
| 用途 | カラー名 | HEX |
|------|---------|-----|
| スライド背景 | BG | #F8FAFC |
| ダークテキスト | DTXT | #1E293B |
| ミディアムテキスト | MTXT | #64748B |
| ライトテキスト | LTXT | #94A3B8 |
| メインネイビー | NAVY | #1E3A5F |
| メインブルー | BLUE | #2563EB |
| 薄ブルー背景 | LBLUE | #EFF6FF |
| ボーダー | BORDER | #CBD5E1 |
| カード背景 | CARD_BG | #F1F5F9 |

### スライド構成パターン

#### タイトルスライド
- 薄グレー背景（#F8FAFC）にグリッド線
- 中央に白カード（枠線付き）
- 角丸タグ（英語ラベル）+ 大タイトル + 区切り線 + サブタイトル
- ツール名をピル型タグで並べる
- フッターに日付・Confidential

#### セクション区切りスライド（PART XX）
- 左1/3：ダークカラー背景＋ドットグリッド装飾＋「PART XX」テキスト
- 右2/3：白背景＋セッションタグ＋大タイトル＋縦ライン＋サブタイトル

#### コンテンツスライド
- 白背景（#FFFFFF）
- **ヘッダー：左縦ライン（BLUE, 0.07インチ幅）＋タイトルテキスト（NAVY, 22pt）**
- 右上にセッション名（小さくLTXT色）
- ヘッダー下に水平ボーダー線
- **KEY MESSAGE ボックス**：薄ブルー背景＋左ブルー縦ライン＋KEY MESSAGEラベル＋本文
- **カードレイアウト**：角がある白カード＋BORDER色枠線
- カードヘッダー：薄ブルー背景＋左アクセントカラー縦ライン
- **下部テイクアウェイバー**：CARD_BG背景＋「今日の持ち帰り」ネイビータグ

### コンポーネント仕様

```python
# ヘッダー
add_rect(slide, 0.25, 0.18, 0.07, 0.56, BLUE)   # 縦ライン
add_text(slide, title, 0.45, 0.2, 9.5, 0.6, size=22, bold=True, color=NAVY)
add_text(slide, session_label, 10.5, 0.28, 2.7, 0.35, size=9, color=LTXT, align=RIGHT)
add_rect(slide, 0, 0.9, 13.33, 0.04, BORDER)    # 下線

# KEY MESSAGEボックス
add_rect(slide, 0.3, t, 12.73, 1.05, LBLUE, line=BLUE, lw=1)
add_rect(slide, 0.3, t, 0.07, 1.05, BLUE)
add_text(slide, "KEY MESSAGE", ...)  # 9pt, BLUE, italic
add_text(slide, title, ...)          # 16pt, bold, NAVY
add_text(slide, body, ...)           # 11pt, MTXT

# カード
add_rect(slide, l, t, w, h, WHITE, line=BORDER, lw=0.8)
add_rect(slide, l, t, w, 0.45, LBLUE, line=BORDER, lw=0.5)   # ヘッダー背景
add_rect(slide, l, t, 0.07, 0.45, accent_color)                # 左アクセント

# ピル型タグ
add_rect(slide, l, t, tw, 0.32, color, rounded=True)  # 角丸=True

# テイクアウェイバー（y=6.9）
add_rect(slide, 0, 6.9, 13.33, 0.6, CARD_BG, line=BORDER, lw=0.5)
add_rect(slide, 0.3, 7.0, 1.8, 0.38, NAVY, rounded=True)
add_text(slide, "今日の持ち帰り", ...)
```

---

## スライド構成テンプレート（推奨16〜18枚）

| # | タイプ | 内容 |
|---|--------|------|
| 1 | タイトル | テーマ・サブタイトル・ツール名タグ |
| 2 | アジェンダ | タイムライン形式・パートとタグ |
| 3 | セクション | PART 01 |
| 4〜5 | コンテンツ | セクション①の詳細（KEY MESSAGE＋カード） |
| 6 | セクション | PART 02 |
| 7〜9 | コンテンツ | セクション②の詳細 |
| 10 | セクション | PART 03 |
| 11〜12 | コンテンツ | セクション③の詳細＋収益シミュレーション等 |
| 13 | セクション | PART 04 |
| 14〜15 | コンテンツ | セクション④の詳細 |
| 16 | コンテンツ | よくある失敗と対策（プラスアルファ） |
| 17 | まとめ | サマリーカード＋今日からできる3つのアクション |

---

## プラスアルファ コンテンツ指針

マインドマップの内容に加えて、以下を必ず補完する：

| 追加要素 | 内容例 |
|---------|-------|
| 統計・数値データ | 市場規模・普及率・収入目安など |
| 収益シミュレーション | フェーズ別の月収モデル表 |
| よくある失敗と対策 | NG例3選＋改善策 |
| 今日からできるアクション | 具体的な3ステップ |
| ツール・サービス名 | 実名を記載してリアリティを出す |

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

XMindが画像（PNG）でしか提供されない場合は、画像を読み込んでテキスト構造を手動で抽出する。

### Step 2: マインドマップ構造の抽出
```python
def extract_topics(node, level=0):
    title = node.get("title", "")
    children = node.get("children", {}).get("attached", [])
    return {"title": title, "level": level,
            "children": [extract_topics(c, level+1) for c in children]}
```

### Step 3: PPTスライドの生成
```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
prs.slide_width  = Inches(13.33)   # ワイド 16:9
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]       # 白紙レイアウト
```

---

## 出力ルール

| フェーズ | 保存先 | ファイル名例 |
|---------|--------|------------|
| 初回生成 | `出力/下書き/` | `下書き_v1.pptx` |
| 修正版 | `出力/下書き/` | `下書き_v2.pptx` |
| 最終承認後 | `出力/最終版/` | `最終版_YYYYMMDD.pptx` |

---

## 作業開始前チェックリスト
- [ ] `読み込み資料/` に元資料（XMind・PNG・テキスト等）が格納されている
- [ ] フォントが **Meiryo UI** に設定されている
- [ ] GenSpark風フォーマット（白背景・カード・縦ライン・セクションスプリット）を使用している
- [ ] プラスアルファのコンテンツ（統計・シミュレーション・失敗対策）が含まれている
- [ ] 「今日の持ち帰り」バーが各コンテンツスライドに付いている
- [ ] スライド数は16〜18枚程度に収める
- [ ] 生成スクリプトは `出力/下書き/generate_ppt_vX.py` に保存する
