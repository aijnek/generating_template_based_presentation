---
name: generating-template-based-presentation
description: テンプレートに完全準拠したスライドを作成するスキル。ユーザーがアップロードしたPPTXテンプレートを分析し、レイアウト構造を理解した上で、python-pptxを使ってテンプレートのデザインを忠実に再現するスライドを生成する。テンプレートの分析→構成決定→レイアウト割当→スライド作成→視覚的検証という5段階ワークフローに従う。'テンプレートを使ってスライドを作って', 'このデザインでプレゼンを作成して', 'テンプレートに合わせて' などのリクエストでトリガーする。テンプレートPPTXファイルが提供された場合は必ずこのスキルを使用する。
---

# Template-Compliant PPTX Skill

テンプレートPPTXファイルを分析し、そのレイアウト・デザインに完全準拠したスライドをpython-pptxで作成するスキル。

## ワークフロー概要

```
Phase 1: テンプレート分析（構造 + 視覚）
Phase 2: スライド構成決定
Phase 3: レイアウト割当
Phase 4: スライド作成
Phase 5: 視覚的検証
Phase 6: ユーザーへの提示
```

**重要**: 必ずこの順序に従うこと。Phase を飛ばさないこと。

---

## Phase 1: テンプレート分析

テンプレートの構造と見た目の両方を把握する。**必ず構造分析と視覚分析の両方を行うこと。**

### 1a. 構造分析（python-pptx）

テンプレートのスライドレイアウトとプレースホルダーを列挙する分析スクリプトを実行する:

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu

prs = Presentation('template.pptx')

# --- スライドレイアウト一覧 ---
print("=" * 60)
print("SLIDE LAYOUTS")
print("=" * 60)
for i, layout in enumerate(prs.slide_layouts):
    print(f'\n--- Layout {i}: "{layout.name}" ---')
    for ph in layout.placeholders:
        phf = ph.placeholder_format
        print(f'  idx={phf.idx}, type={phf.type}, name="{ph.name}"')
        print(f'    position: left={ph.left}, top={ph.top}')
        print(f'    size: width={ph.width}, height={ph.height}')
        if ph.has_text_frame:
            for para in ph.text_frame.paragraphs:
                if para.text.strip():
                    print(f'    default_text: "{para.text[:50]}..."')

# --- 既存スライド一覧（テンプレートにサンプルスライドがある場合）---
print("\n" + "=" * 60)
print("EXISTING SLIDES")
print("=" * 60)
for i, slide in enumerate(prs.slides):
    layout = slide.slide_layout
    print(f'\nSlide {i}: layout="{layout.name}"')
    for shape in slide.shapes:
        print(f'  shape: name="{shape.name}", type={shape.shape_type}')
        if shape.is_placeholder:
            phf = shape.placeholder_format
            print(f'    placeholder: idx={phf.idx}, type={phf.type}')
        if shape.has_text_frame:
            text = shape.text_frame.text[:80]
            if text.strip():
                print(f'    text: "{text}"')
        print(f'    pos: left={shape.left}, top={shape.top}, w={shape.width}, h={shape.height}')
```

**分析で確認すべきポイント:**
- 各レイアウトのプレースホルダーの `idx` 値（アクセスに必須）
- プレースホルダーの `type`（TITLE, BODY, PICTURE, TABLE, CHART など）
- 位置とサイズ（重なり回避のため）
- プレースホルダー以外のシェイプ（装飾要素、ロゴなど）

### 1b. 視覚分析（PDF化）

テンプレートをPDFに変換し、各ページの見た目を視覚的に確認する:

```bash
# テンプレートをPDF化
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf template.pptx

# 各ページを画像化
pdftoppm -jpeg -r 150 template.pdf template-slide
```

生成された画像（`template-slide-01.jpg`, `template-slide-02.jpg`, ...）を確認し、各レイアウトの**視覚的な役割**を判断する:
- タイトルスライド
- セクション区切り
- コンテンツスライド（テキスト中心）
- 図表スライド
- 画像配置スライド
- まとめ/エンディング

### 1c. 分析結果の整理

Phase 1の結果を以下の形式で整理する:

```
レイアウト0: "Title Slide" → タイトル(idx=0) + サブタイトル(idx=1) → 表紙用
レイアウト1: "Title and Content" → タイトル(idx=0) + 本文(idx=1) → 一般コンテンツ用
レイアウト2: "Two Content" → タイトル(idx=0) + 左本文(idx=1) + 右本文(idx=2) → 比較・2列用
...
```

この整理結果をPhase 3で参照する。

---

## Phase 2: スライド構成決定

ユーザーの要件に基づき、スライドの内容と構成を決定する。

**確認事項:**
- プレゼンの目的（報告、提案、教育、etc.）
- 対象者
- スライド枚数の目安
- 各スライドの内容（タイトル、本文、図表の有無）
- 使用言語

**出力:** スライドごとの内容案（タイトル・本文・メモ）

---

## Phase 3: レイアウト割当

Phase 1の分析結果とPhase 2の構成をマッチングし、各スライドにテンプレートのレイアウトを割り当てる。

**原則:**
- コンテンツの種類に最も適したレイアウトを選ぶ
- **同じレイアウトの連続使用は最大2〜3回まで**（単調にならないよう意識する）
- テンプレートに存在するレイアウトのみを使用する（自分でデザインしない）
- プレースホルダーの数と種類がコンテンツに合っているか確認する

**出力例:**
```
Slide 0: Layout 0 ("Title Slide") → 表紙
Slide 1: Layout 1 ("Title and Content") → 背景・課題説明
Slide 2: Layout 2 ("Two Content") → 現状と目標の比較
Slide 3: Layout 1 ("Title and Content") → 提案内容
Slide 4: Layout 5 ("Blank") → カスタム図表
...
```

---

## Phase 4: スライド作成

python-pptxを使ってスライドを作成する。**参考資料は `references/` ディレクトリにある。**

### 基本構造

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData

# テンプレートを読み込み
prs = Presentation('template.pptx')

# テンプレートの既存スライドを削除（テンプレートのサンプルページ）
# 注意: 逆順で削除する
for i in range(len(prs.slides) - 1, -1, -1):
    remove_slide(prs, i)

# 各スライドを作成
# Layout index は Phase 1 の分析結果に基づく
```

### スライド削除ヘルパー

```python
def remove_slide(prs, slide_index):
    """スライドをプレゼンテーションから削除する"""
    sldIdLst = prs._element.sldIdLst
    target_sldId = sldIdLst.sldId_lst[slide_index]
    sldIdLst.remove(target_sldId)
    prs.part.drop_rel(target_sldId.rId)
```

### プレースホルダーへのテキスト挿入

```python
layout = prs.slide_layouts[LAYOUT_INDEX]
slide = prs.slides.add_slide(layout)

# タイトル（idx=0 が一般的）
title_ph = slide.placeholders[0]
title_ph.text = "スライドタイトル"

# 本文（idx=1 が一般的）
body_ph = slide.placeholders[1]
tf = body_ph.text_frame
tf.clear()

# 最初の段落
p = tf.paragraphs[0]
p.text = "最初のポイント"
p.alignment = PP_ALIGN.LEFT

# 追加の段落
p = tf.add_paragraph()
p.text = "2番目のポイント"
p.alignment = PP_ALIGN.LEFT
p.level = 0  # インデントレベル（0=トップ, 1=サブ）
```

### 文字書式の設定

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor

run = p.add_run()
run.text = "書式付きテキスト"
font = run.font
font.name = "メイリオ"       # 日本語フォント
font.size = Pt(14)
font.bold = True
font.color.rgb = RGBColor(0x33, 0x33, 0x33)
```

### テキスト配置の注意点

プレースホルダーのテキストを書き換える際のルール:

1. **`text_frame.clear()` を使う**: 既存テキストを安全に消去
2. **最初の段落は `paragraphs[0]`**: clear後も1つ残る
3. **追加は `add_paragraph()`**: 2番目以降の段落
4. **書式はrunレベルで設定**: `p.text =` はショートカット。細かい書式が必要な場合は `run` を使う
5. **テンプレートの書式を継承させたい場合**: font属性を `None` のままにする（明示的に設定すると継承が切れる）

### 画像の挿入

```python
# ピクチャープレースホルダーへの挿入
pic_ph = slide.placeholders[PIC_IDX]  # Phase 1で確認したidx
picture = pic_ph.insert_picture('image.png')
# 注意: insert_picture後は元のpic_phは無効になる

# 任意位置への画像追加
from pptx.util import Inches
slide.shapes.add_picture(
    'image.png',
    left=Inches(1), top=Inches(2),
    width=Inches(4), height=Inches(3)
)
```

### テーブルの作成

```python
# テーブルプレースホルダーへの挿入
table_ph = slide.placeholders[TABLE_IDX]
graphic_frame = table_ph.insert_table(rows=4, cols=3)
table = graphic_frame.table

# または任意位置に追加
shape = slide.shapes.add_table(
    rows=4, cols=3,
    left=Inches(1), top=Inches(2),
    width=Inches(8), height=Inches(3)
)
table = shape.table

# セルへのデータ入力
table.cell(0, 0).text = "ヘッダー1"
table.cell(0, 1).text = "ヘッダー2"
table.cell(1, 0).text = "データ1"

# セル書式
from pptx.util import Pt
cell = table.cell(0, 0)
for paragraph in cell.text_frame.paragraphs:
    for run in paragraph.runs:
        run.font.bold = True
        run.font.size = Pt(12)
```

### チャートの作成

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

chart_data = CategoryChartData()
chart_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
chart_data.add_series('売上', (120, 150, 180, 200))
chart_data.add_series('利益', (30, 45, 55, 70))

# チャートプレースホルダーへの挿入
chart_ph = slide.placeholders[CHART_IDX]
graphic_frame = chart_ph.insert_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, chart_data
)
chart = graphic_frame.chart

# または任意位置に追加
x, y, cx, cy = Inches(1), Inches(2), Inches(8), Inches(4)
graphic_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
)
chart = graphic_frame.chart

# チャートのカスタマイズ
chart.has_legend = True
chart.legend.include_in_layout = False
```

### プレースホルダー以外のシェイプ操作

テンプレートに装飾的なシェイプ（ロゴ、区切り線など）がある場合、それらはそのまま残す。テキストを含むシェイプのみ必要に応じて編集する:

```python
for shape in slide.shapes:
    if not shape.is_placeholder and shape.has_text_frame:
        # 必要に応じてテキストを更新
        if "XXXX" in shape.text or "Lorem" in shape.text:
            shape.text_frame.text = "置換テキスト"
```

### スピーカーノートの追加

```python
notes_slide = slide.notes_slide
notes_text_frame = notes_slide.notes_text_frame
notes_text_frame.text = "このスライドの説明メモ"
```

### よくある落とし穴と対策

| 問題 | 原因 | 対策 |
|------|------|------|
| プレースホルダーが見つからない | idx値の誤り | Phase 1の分析結果を再確認 |
| テキストが溢れる | コンテンツが長すぎる | フォントサイズ縮小 or テキスト分割 |
| 書式が崩れる | font属性の明示設定 | 継承させたい属性は設定しない（None） |
| 画像挿入後にエラー | insert_picture後の参照無効化 | 戻り値を使う |
| 日本語が文字化け | フォント未指定 | "メイリオ", "游ゴシック" 等を明示 |
| テンプレートのサンプルテキストが残る | 削除漏れ | Phase 5でgrep検証 |

---

## Phase 5: 視覚的検証（必須）

**作成したスライドは必ず視覚的に検証すること。** テキストの重なり、はみ出し、レイアウト崩れはコードだけでは検出できない。

### 5a. PDF変換・画像化

```bash
# PPTXをPDFに変換
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx

# PDFを各ページの画像に変換
pdftoppm -jpeg -r 150 output.pdf slide
```

### 5b. コンテンツ検証

```bash
# テキスト抽出で内容確認
pip install "markitdown[pptx]" --break-system-packages -q
python -m markitdown output.pptx

# テンプレートの残留テキスト検出
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|click.*(add|to)|placeholder|sample"
```

### 5c. 視覚的検証

生成された画像（`slide-01.jpg`, `slide-02.jpg`, ...）を1枚ずつ確認する。

**チェック項目:**
- [ ] テキストの重なり・はみ出しがない
- [ ] フォントサイズが適切（小さすぎない）
- [ ] テンプレートのデザイン要素（ロゴ、装飾）が維持されている
- [ ] テンプレートのサンプルテキストが残っていない
- [ ] 余白が適切（端に詰まりすぎていない）
- [ ] 日本語テキストが正しく表示されている
- [ ] チャート・テーブルが正しく描画されている
- [ ] 画像が適切なサイズで配置されている

### 5d. 修正ループ

問題が見つかった場合:
1. 問題を特定（どのスライドの何が問題か）
2. コードを修正
3. 再生成 → 再変換 → 再検証
4. 問題が解消されるまで繰り返す

**少なくとも1回は修正サイクルを回すこと。** 初回で完璧なことはまずない。

---

## Phase 6: ユーザーへの提示

完成したPPTXファイルを `/mnt/user-data/outputs/` にコピーし、`present_files` ツールで提示する。

検証で使用したPDF画像も合わせて提示すると、ユーザーがダウンロード前に内容を確認できて親切。

---

## 依存関係

```bash
pip install python-pptx --break-system-packages
pip install "markitdown[pptx]" --break-system-packages
# LibreOffice (soffice) と Poppler (pdftoppm) はプリインストール済み
```

---

## 参考資料

詳細なAPIリファレンスは以下を参照:

- `references/python-pptx-concepts.md` — Presentation, Slide, Layout, Placeholder, Shape, Text の操作方法
- `references/python-pptx-charts-tables.md` — Chart, Table の作成・操作方法

**必ず作成開始前に参考資料を読むこと。** 特にプレースホルダーのidxアクセス、insert後の参照無効化、テキスト階層（TextFrame → Paragraph → Run）は重要。