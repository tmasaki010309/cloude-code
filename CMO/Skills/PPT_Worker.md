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

### カラーパレット（青系統一）
色は青系5段階で統一する（緑・オレンジ・赤・紫は使用しない）。

| 用途 | カラー名 | HEX |
|------|---------|-----|
| スライド背景 | BG | #F8FAFC |
| ダークテキスト | DTXT | #1E293B |
| ミディアムテキスト | MTXT | #475569 |
| ライトテキスト | LTXT | #94A3B8 |
| 最濃ネイビー | NAVY | #0F1F3D |
| 濃ネイビー | B1 | #1E3A5F |
| ダークブルー | B2 | #1D4ED8 |
| メインブルー | B3 | #2563EB |
| ブライトブルー | B4 | #3B82F6 |
| 超薄ブルー背景 | LBLUE | #EFF6FF |
| ボーダー（薄ブルー） | BORDER | #BFDBFE |
| カード背景 | CARD | #F1F5F9 |

セクション区切りスライド（PART 1〜4）は以下の濃淡で区別する：
- Part1: #0F1F3D（最濃）、Part2: #1A3560、Part3: #1A3A8F、Part4: #1D4ED8

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
# ヘッダー（高さ1.0インチ）
add_rect(slide, 0.25, 0.2, 0.07, 0.6, BLUE)             # 縦ライン
add_text(slide, title, 0.45, 0.2, 9.5, 0.65, size=22, bold=True, color=NAVY)
add_text(slide, session_label, 10.2, 0.28, 3.0, 0.4, size=11, color=LTXT, align=RIGHT)
add_rect(slide, 0, 1.0, 13.33, 0.04, BORDER)             # 下線

# KEY MESSAGEボックス（高さ1.35インチ）
add_rect(slide, 0.3, t, 12.73, 1.35, LBLUE, line=BLUE, lw=1)
add_rect(slide, 0.3, t, 0.07, 1.35, BLUE)
add_text(slide, "KEY MESSAGE", ...)  # 11pt, BLUE, italic
add_text(slide, title, ...)          # 17pt, bold, NAVY
add_text(slide, body, ...)           # 12pt, MTXT

# カード
add_rect(slide, l, t, w, h, WHITE, line=BORDER, lw=0.8)
add_rect(slide, l, t, w, 0.5, LBLUE, line=BORDER, lw=0.5)    # ヘッダー背景
add_rect(slide, l, t, 0.07, 0.5, accent_color)                 # 左アクセント

# ピル型タグ（動的幅：max(len(text)*0.18+0.35, 1.1)で計算）
tw = max(len(text)*0.18+0.35, 1.1)
add_rect(slide, l, t, tw, 0.34, color, rounded=True)  # 角丸=True
# ※右端13.1インチ超になる場合は描画しない（はみ出し防止）

# テイクアウェイバー（y=6.85）
add_rect(slide, 0, 6.85, 13.33, 0.65, CARD, line=BORDER, lw=0.5)
add_rect(slide, 0.3, 6.97, 1.9, 0.42, NAVY, rounded=True)
add_text(slide, "今日の持ち帰り", ..., size=11, bold=True)
# アイテムテキスト: size=12, x=2.4から5.4インチ間隔
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

## フォントサイズ規則（視認性ルール）

| 用途 | 最小サイズ | 推奨 |
|------|----------|------|
| スライドタイトル | 22pt | 22〜28pt |
| セクションタイトル | 40pt | 44〜52pt |
| カードタイトル | 13pt | 14〜16pt |
| KEY MESSAGE タイトル | 16pt | 17〜18pt |
| KEY MESSAGE ラベル | 11pt | 11pt（italic） |
| KEY MESSAGE 本文 | 12pt | 12〜13pt |
| 本文・箇条書き（Level 0） | 12pt | 13pt |
| 本文・箇条書き（Level 1以下） | 12pt | 12pt（Level上限で削減しない） |
| タグ・ラベル | 11pt | 11pt |
| テイクアウェイ項目 | 12pt | 12pt |
| セッション名（右上） | 11pt | 11pt |

**グレーテキスト（MTXT）を使う場合は必ず太字（bold=True）を適用し視認性を確保する。**

---

## 作業開始前チェックリスト
- [ ] `読み込み資料/` に元資料（XMind・PNG・テキスト等）が格納されている
- [ ] フォントが **Meiryo UI** に設定されている
- [ ] カラーが **青系5段階** に統一されている（緑・オレンジ・赤・紫不使用）
- [ ] GenSpark風フォーマット（白背景・カード・縦ライン・セクションスプリット）を使用している
- [ ] プラスアルファのコンテンツ（統計・シミュレーション・失敗対策）が含まれている
- [ ] 各カードの根拠・詳細は **5項目以上** 記載している
- [ ] 内容が多いページは **1ページに詰め込まず分割** する
- [ ] 「今日の持ち帰り」バーが各コンテンツスライドに付いている
- [ ] すべてのテキストが **最小12pt** を満たしている
- [ ] グレーテキスト使用時は **太字** で視認性を確保している
- [ ] スライド数は16〜18枚程度に収める
- [ ] 生成スクリプトは `出力/下書き/generate_ppt_vX.py` に保存する
