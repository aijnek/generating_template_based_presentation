---
name: generating-template-based-presentation
description: ユーザーがアップロードしたPowerPointテンプレートに完全準拠したスライドを作成します。テンプレートのレイアウト、色、フォント、装飾要素を正確に保持しながら、新しいコンテンツでプレゼンテーションを作成します。
---

# Template-Conformant PPTX Creation Skill

このスキルは、ユーザー提供のPowerPointテンプレートに完全に準拠した新規プレゼンテーションを作成します。

## コア原則

1. **テンプレートの完全な尊重** - ユーザーのテンプレートのデザイン、色、フォント、レイアウトを100%保持
2. **テキストクリアの徹底** - スライドコピー後、必ず全テキストをクリアしてから新しい内容を設定（重なり防止）
3. **レイアウトの賢い選択** - 各スライドのコンテンツに最適なレイアウトを選択
4. **デザインの多様性** - 同じレイアウトの連続を避け、視覚的なリズムを作る
5. **視覚的確認の徹底** - 全スライドをPDF化・画像化して目視確認。特に装飾との重なりをチェック
6. **品質保証の徹底** - 文字の重なり、はみ出し、不具合を必ずチェック

## ワークフロー

### ステップ1: 要件ヒアリング

ユーザーからの情報収集：

```
必須情報:
- プレゼンテーションの目的・テーマ
- 対象オーディエンス
- 各スライドの内容（箇条書きまたは詳細）

任意情報:
- 希望するスライド枚数
- 特に強調したいポイント
- 避けたいレイアウトやスタイル
```

**ユーザーへの確認:**
- テンプレートファイルがアップロード済みか確認
- コンテンツの詳細度を確認（概要のみ vs 詳細テキスト）

---

### ステップ2: テンプレート分析

#### 2a. テンプレートの視覚確認

```bash
# テンプレートのサムネイルを生成
python scripts/thumbnail.py template.pptx
```

各スライドを視覚的に確認し、以下を把握：
- スライド数
- 全体的なデザインスタイル
- 色使いとフォント

#### 2b. テンプレート構造の詳細分析

```bash
# テンプレートをテキスト形式で抽出
python -m markitdown template.pptx > template_content.txt
```

各スライドについて記録：
- レイアウト名（推定）
- 主要な構成要素
- テキストボックスの配置
- 画像やグラフの配置

#### 2c. レイアウトカタログの作成

テンプレート内の各スライドを分類：

| スライド番号 | レイアウトタイプ | 特徴 | 適した用途 |
|------------|----------------|------|----------|
| 1 | タイトルスライド | 大きなタイトル、サブタイトル | 表紙、セクション区切り |
| 2 | 箇条書き | タイトル + 箇条書きエリア | リスト、要点整理 |
| 3 | 2カラム | 左右分割レイアウト | 比較、before/after |
| ... | ... | ... | ... |

**ユーザーへの提示:**
- サムネイル画像と共にレイアウトカタログを提示
- 各レイアウトの推奨用途を説明

---

### ステップ3: コンテンツ構成設計

#### 3a. スライドリストの作成

ユーザーの要件に基づき、スライド構成案を作成：

```
スライド1: [タイトル] - レイアウト: スライド1（タイトルスライド）
スライド2: [背景説明] - レイアウト: スライド2（箇条書き）
スライド3: [課題] - レイアウト: スライド5（強調スライド）
スライド4: [解決策] - レイアウト: スライド3（2カラム）
...
```

#### 3b. レイアウトマッチングの基準

各コンテンツタイプに対するレイアウト選択：

**タイトル・セクション区切り:**
- 大きなタイトルエリアを持つスライド
- 装飾的で視覚的にインパクトのあるデザイン

**箇条書き・リスト:**
- タイトル + テキストエリアの構成
- 十分な余白のあるレイアウト

**比較・対比:**
- 2カラムレイアウト
- 左右または上下に分割された構成

**図表・データ中心:**
- 画像やグラフ用の大きなエリアを持つレイアウト
- テキストは最小限

**複数トピック:**
- 3〜4分割のグリッドレイアウト
- 各セクションが独立したボックス

#### 3c. デザイン多様性の確保

**避けるべきパターン:**
- 同じレイアウトを3回以上連続使用
- 全スライドで箇条書きのみ
- 視覚的な変化のない単調な構成

**推奨パターン:**
- タイトル → コンテンツA → コンテンツB → セクション区切り → ...
- レイアウトを2〜3スライドごとに変える
- 視覚的に「重い」スライドと「軽い」スライドを交互に

**ユーザーへの確認:**
- スライド構成案とレイアウト選択を提示
- フィードバックを収集し、必要に応じて修正

---

### ステップ4: スライド作成

#### 4a. テンプレートの準備

```bash
# テンプレートファイルをコピー
cp template.pptx working_presentation.pptx
```

#### 4b. Python-pptxを使った編集

```python
from pptx import Presentation

# テンプレートを読み込み
prs = Presentation('template.pptx')

# 新規スライドを作成する方針:
# 1. テンプレート内の適切なスライドを特定
# 2. そのスライドをコピー（duplicate_slide関数を使用）
# 3. コピーしたスライドのコンテンツを置き換え
```

#### 4c. コンテンツ置き換えの手順

各スライドについて：

1. **テンプレートから適切なスライドを複製**
   ```python
   # 選択したレイアウトに最も近いテンプレートスライドを特定
   reference_slide_idx = template_slides[layout_choice]
   new_slide = duplicate_slide(prs, reference_slide_idx)
   ```

2. **【重要】まず全テキストをクリア**
   ```python
   # テンプレートの元テキストを完全に削除（重なり防止のため必須）
   for shape in new_slide.shapes:
       if shape.has_text_frame:
           shape.text_frame.clear()
   ```
   
   **注意：** このステップを省略すると、テンプレートの元テキストが残り、新しいテキストと重なって表示される問題が発生します。必ず実行してください。

3. **テキストの置き換え**
   ```python
   for shape in new_slide.shapes:
       if shape.has_text_frame:
           # タイトル、本文などを識別して置き換え
           if is_title_shape(shape):
               shape.text = new_title
           elif is_content_shape(shape):
               shape.text = new_content
   ```

4. **プレースホルダーの尊重**
   - テンプレートのプレースホルダー構造を保持
   - テキストのフォーマット（フォント、サイズ、色）を維持

5. **装飾要素の保持**
   - 背景画像、ロゴ、装飾図形はそのまま保持
   - これらを誤って削除・変更しない

#### 4d. スライド順序の調整

```python
# 新規作成したスライドを適切な順序に並べる
# 元のテンプレートスライドで不要なものを削除
```

---

### ステップ5: 品質保証

#### 5a. テンプレート準拠チェック

**チェック項目:**
- [ ] 色スキームが一貫している（テンプレートの色を使用）
- [ ] フォントが一貫している（テンプレートのフォントを使用）
- [ ] 背景・装飾要素が正しく保持されている
- [ ] レイアウト構造が崩れていない

**確認方法:**
```bash
# 作成したスライドを視覚確認
python scripts/thumbnail.py working_presentation.pptx

# テンプレートと比較
# 同じスライド番号を見比べて、デザインが保持されているか確認
```

#### 5b. コンテンツ品質チェック

**【重要】視覚的確認（必須）:**

文字の重なりや装飾との干渉は、テキスト抽出だけでは検出できません。必ず各スライドを画像として視覚的に確認してください。

```bash
# ステップ1: PDFに変換
libreoffice --headless --convert-to pdf --outdir . working_presentation.pptx

# ステップ2: 各スライドを画像として抽出（150dpiで十分）
# 全スライドを確認
pdftoppm working_presentation.pdf slide -png -r 150

# または特定のページのみ確認（例：2ページ目）
pdftoppm working_presentation.pdf slide -png -f 2 -l 2 -r 150

# ステップ3: 生成された画像を確認
ls -lh slide*.png
```

**画像を使った視覚的確認:**
```python
# viewツールで各スライドを表示
view('/path/to/slide-01.png')
view('/path/to/slide-02.png')
# ... 全スライドを確認
```

**チェックポイント:**
- [ ] タイトルがテキストボックスに収まっている
- [ ] 箇条書きの全項目が見える（下部の装飾と重なっていない）
- [ ] テキストが背景や装飾図形と重なっていない
- [ ] 文字色と背景色のコントラストが十分
- [ ] 長い文章がテキストボックスからはみ出していない
- [ ] 特に下部の装飾エリア（帯、バー、グラデーション）との重なりに注意

**問題発見時の対処法:**

1. **箇条書きが装飾と重なる場合:**
   ```python
   # 項目数を減らす（最も効果的）
   # 元のテンプレートの項目数を確認し、それを超えない
   
   # または項目を統合
   # 例: "概要" "目的" → "概要と目的"
   ```

2. **フォントサイズの調整:**
   ```python
   # 全体的にサイズを縮小
   for shape in slide.shapes:
       if shape.has_text_frame:
           for paragraph in shape.text_frame.paragraphs:
               for run in paragraph.runs:
                   if run.font.size and run.font.size > Pt(20):
                       run.font.size = Pt(18)  # 2ptずつ縮小
   ```

3. **テキストボックスの高さ制限:**
   ```python
   # 装飾エリアより上に制限
   from pptx.util import Inches
   
   for shape in slide.shapes:
       if shape.has_text_frame and not is_title_shape(shape):
           # 下部の装飾が5.5インチから始まる場合
           max_bottom = Inches(5.2)  # 安全マージン
           if shape.top + shape.height > max_bottom:
               # テキストボックスを装飾より上に制限
               shape.height = max_bottom - shape.top
   ```

**修正後の再確認:**
```bash
# 修正したらもう一度PDFと画像を生成して確認
libreoffice --headless --convert-to pdf --outdir . working_presentation.pptx
pdftoppm working_presentation.pdf verified -png -r 150

# 修正箇所を視覚的に確認
view('verified-02.png')  # 修正したスライドを確認
```

#### 5c. デザイン多様性チェック

レイアウトの使用状況を確認：
```
スライド1-3: レイアウトA
スライド4-5: レイアウトB  ← OK
スライド6-9: レイアウトA  ← 問題：4回連続で同じレイアウト
```

**単調さの兆候:**
- 同じレイアウトが3回以上連続
- すべてのスライドが箇条書きのみ
- 視覚的な変化がない

**改善方法:**
- 一部のスライドのレイアウトを変更
- コンテンツを統合または分割してレイアウトを多様化

#### 5d. 最終調整

不具合が見つかった場合：

1. **文字の重なり** → フォントサイズを縮小、行間を調整
2. **テキストのはみ出し** → テキストボックスを拡大、または内容を簡潔化
3. **レイアウト崩れ** → 該当スライドを再作成
4. **色・フォントの不一致** → テンプレートの設定を再適用

#### 5e. 視覚的確認（必須）

**このステップは省略不可。全スライドを必ず目視確認すること。**

```bash
# 1. PDFに変換
libreoffice --headless --convert-to pdf --outdir . working_presentation.pptx

# 2. 全スライドを画像化
pdftoppm working_presentation.pdf slide -png -r 150

# 3. 各スライドを確認
ls -lh slide*.png
```

**Pythonで自動的に全スライドを確認:**
```python
import os
import glob

# 画像ファイルを取得
png_files = sorted(glob.glob('slide-*.png'))

print(f"生成された画像: {len(png_files)}枚")

# viewツールで各スライドを表示して確認
for png_file in png_files:
    print(f"\n確認中: {png_file}")
    # viewツールで表示
    # view(png_file)
    # 問題があれば記録
```

**確認項目（各スライドごと）:**
- [ ] タイトルが完全に表示されているか
- [ ] 箇条書きの全項目が見えるか（特に最後の項目）
- [ ] 下部の装飾（帯、バー）とテキストが重なっていないか
- [ ] テキストボックスからはみ出していないか
- [ ] 背景や装飾図形との重なりがないか
- [ ] 全体的なバランスが良いか

**問題が見つかった場合:**

1. 問題のあるスライド番号を記録
2. 原因を特定（項目数過多、テキスト長すぎ、など）
3. 5a-5dの手順に戻って修正
4. 修正後、再度視覚的確認を実施（5eを繰り返す）

**修正例:**
```python
# 例: スライド2で箇条書きが装飾と重なっている場合
slide2 = prs.slides[1]

# 原因: 7項目あるが、テンプレートは5項目想定
# 解決策: 項目を5つに統合
# 修正後、もう一度PDF化して確認
```

---

### ステップ6: 不要なスライドの削除

テンプレートに含まれていた元のスライドで、新規作成したスライドに置き換えられていないものを削除：

```python
# 使用しなかったテンプレートスライドを削除
slides_to_keep = [新規作成したスライドのインデックス]
slides_to_remove = [使わなかったテンプレートスライドのインデックス]

for idx in reversed(slides_to_remove):
    rId = prs.slides._sldIdLst[idx].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[idx]
```

**確認:**
- 削除後、スライド番号と内容が意図通りか確認
- 必要なスライドを誤って削除していないか確認

---

### ステップ7: 最終確認とユーザー提示

#### 7a. 変更履歴の報告

ユーザーに以下を報告：

```
完成レポート:
- 作成したスライド数: X枚
- 使用したレイアウトの種類: Y種類
- 実施した調整:
  * スライド3: タイトルのフォントサイズを24pt→22ptに調整
  * スライド7: テキストボックスを下に5mm移動
- 削除したテンプレートスライド: Z枚
```

#### 7b. QAサマリー

品質保証の結果：

```
✓ テンプレート準拠: 色・フォント・レイアウトすべて維持
✓ 文字の重なり: なし
✓ テキストのはみ出し: なし  
✓ デザイン多様性: Xパターンのレイアウトを使用、単調さなし
```

#### 7c. ファイル提示

```bash
# 最終ファイルを出力ディレクトリにコピー
cp working_presentation.pptx /mnt/user-data/outputs/final_presentation.pptx
```

**ユーザーへの提示:**
- 完成したファイルをダウンロード可能にする
- サムネイル画像で全体像を見せる
- 必要に応じて修正対応

---

## ヘルパー関数

### duplicate_slide()

```python
def duplicate_slide(prs, slide_idx):
    """指定されたスライドを複製"""
    from pptx.util import Inches
    from copy import deepcopy
    
    source = prs.slides[slide_idx]
    
    # 同じレイアウトで新規スライドを作成
    blank_slide = prs.slides.add_slide(source.slide_layout)
    
    # すべての要素をコピー
    for shape in source.shapes:
        el = shape.element
        newel = deepcopy(el)
        blank_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
    
    return blank_slide
```

### is_title_shape()

```python
def is_title_shape(shape):
    """シェイプがタイトルかどうかを判定"""
    if not shape.has_text_frame:
        return False
        
    if shape.is_placeholder:
        ph_type = shape.placeholder_format.type
        # PP_PLACEHOLDER.TITLE (1) or SUBTITLE (3) or CENTER_TITLE (14)
        return ph_type in [1, 3, 14]
    
    # プレースホルダーでない場合、位置とサイズで判定
    # 上部にある大きなテキストボックス
    if shape.top < Inches(2) and shape.height > Inches(0.5):
        return True
    
    return False
```

### is_content_shape()

```python
def is_content_shape(shape):
    """シェイプがコンテンツ（本文）かどうかを判定"""
    if not shape.has_text_frame:
        return False
        
    if shape.is_placeholder:
        ph_type = shape.placeholder_format.type
        # BODY (2) or OBJECT (7)
        return ph_type in [2, 7]
    
    # タイトルでない大きなテキストボックス
    if not is_title_shape(shape) and shape.height > Inches(1):
        return True
    
    return False
```

### create_slide()

```python
def create_slide(prs, template_idx, title, content=None, is_bullet=False):
    """テンプレートからスライドを作成"""
    # スライドを複製
    new_slide = duplicate_slide(prs, template_idx)
    
    # 【重要】まず全テキストシェイプの元テキストをクリア
    # このステップを忘れると、テンプレートテキストと新テキストが重なる
    for shape in new_slide.shapes:
        if shape.has_text_frame:
            shape.text_frame.clear()
    
    # タイトルを設定
    for shape in new_slide.shapes:
        if is_title_shape(shape):
            shape.text = title
            break
    
    # コンテンツを設定
    if content:
        for shape in new_slide.shapes:
            if is_content_shape(shape):
                if is_bullet and isinstance(content, list):
                    # 箇条書きの場合
                    text_frame = shape.text_frame
                    for i, item in enumerate(content):
                        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
                        p.text = item
                        p.level = 0
                else:
                    # 通常のテキスト
                    shape.text = content
                break
    
    return new_slide
```

### analyze_layout_usage()

```python
def analyze_layout_usage(prs):
    """レイアウトの使用状況を分析"""
    layout_usage = {}
    consecutive_count = {}
    prev_layout = None
    
    for slide in prs.slides:
        layout_name = slide.slide_layout.name
        
        # カウント
        layout_usage[layout_name] = layout_usage.get(layout_name, 0) + 1
        
        # 連続使用のチェック
        if layout_name == prev_layout:
            consecutive_count[layout_name] = consecutive_count.get(layout_name, 1) + 1
        else:
            prev_layout = layout_name
    
    # 問題のある連続使用を報告
    issues = []
    for layout, count in consecutive_count.items():
        if count >= 3:
            issues.append(f"{layout}が{count}回連続使用されています")
    
    return layout_usage, issues
```

---

## トラブルシューティング

### 問題: 箇条書きが下部の装飾と重なる【重要】

**症状:** アジェンダなどの箇条書きスライドで、最後の数項目が下部の装飾（帯、バー、グラデーション）と重なって読めなくなる

**原因:** 元のテンプレートの項目数より多くの項目を追加している

**解決策（優先順）:**

1. **項目数を減らす（最も効果的）:**
   ```python
   # 元のテンプレートを確認
   # テンプレートが5項目なら、作成するスライドも5項目以下に
   
   # 項目を統合
   # 例: 
   # "Agent Skillsとは", "なぜ必要か" → "Agent Skillsとは・なぜ必要か"
   # "構造", "Progressive Disclosure" → "構造とProgressive Disclosure"
   ```

2. **項目を複数スライドに分割:**
   ```python
   # 8項目ある場合
   # スライド1: アジェンダ（前半） - 4項目
   # スライド2: アジェンダ（後半） - 4項目
   ```

3. **フォントサイズを縮小:**
   ```python
   for shape in slide.shapes:
       if shape.has_text_frame and not is_title_shape(shape):
           for paragraph in shape.text_frame.paragraphs:
               for run in paragraph.runs:
                   if run.font.size and run.font.size > Pt(18):
                       run.font.size = Pt(18)
   ```

4. **テキストボックスの高さを制限（最終手段）:**
   ```python
   # 装飾エリアを特定し、その上までに制限
   from pptx.util import Inches
   
   for shape in slide.shapes:
       if shape.has_text_frame and not is_title_shape(shape):
           # 装飾が始まる位置より上に制限
           max_bottom = Inches(5.2)  # 装飾の開始位置 - 0.3インチ
           if shape.top + shape.height > max_bottom:
               shape.height = max_bottom - shape.top
   ```

**予防策:**
- スライド作成前に元のテンプレートの項目数を確認
- 視覚的確認を必ず実施（PDF→画像変換）
- 疑わしい場合は少なめの項目数でスタート

### 問題: テンプレートのテキストと新しいテキストが重なる【最重要】

**症状:** スライドに元のテンプレートテキストと新しいテキストの両方が表示され、重なって見える

**原因:** duplicate_slide()でスライドをコピーした後、元のテキストをクリアせずに新しいテキストを設定している

**解決策:**
```python
# スライドコピー後、必ず最初に全テキストをクリア
new_slide = duplicate_slide(prs, template_idx)

# 【必須】全テキストシェイプをクリア
for shape in new_slide.shapes:
    if shape.has_text_frame:
        shape.text_frame.clear()

# この後で新しいテキストを設定
for shape in new_slide.shapes:
    if is_title_shape(shape):
        shape.text = new_title
```

**予防策:**
- create_slide()関数の最初のステップとして、必ずテキストクリアを実行
- コード作成時、このステップを忘れないようコメントで明記

### 問題: テンプレートのレイアウトが少なすぎる

**症状:** テンプレートに2-3種類のスライドしかない

**解決策:**
1. 同じレイアウトでも配置を工夫して変化をつける
2. テキストボックスの追加・削除で視覚的な変化を作る
3. ユーザーに追加のテンプレートスライドの作成を提案

### 問題: テキストがテキストボックスに収まらない

**症状:** 長いテキストがはみ出す、見切れる

**解決策:**
1. フォントサイズを段階的に縮小（最小12pt）
2. 行間を調整（line_spacingを0.9〜1.2に）
3. テキストボックスのサイズを拡大
4. コンテンツを2つのスライドに分割

### 問題: 装飾要素が消える・ずれる

**症状:** ロゴ、背景画像、装飾図形が正しく表示されない

**解決策:**
1. duplicate_slide()を使ってスライド全体をコピー
2. テキストのみを置き換え、他の要素には触れない
3. z-orderを確認（装飾要素が背面にあるか）

### 問題: フォント・色が一致しない

**症状:** テンプレートと異なるフォントや色が使われる

**解決策:**
1. テンプレートのtheme colorsを確認
2. テキスト置き換え時、既存のフォーマットを保持
3. 新規テキストには明示的にスタイルを適用

---

## ベストプラクティス

1. **【最重要】スライドコピー後は必ずテキストをクリア** - duplicate_slide()の直後に、全テキストシェイプをclear()することで、テンプレートテキストとの重なりを防ぐ

2. **【重要】全スライドの視覚的確認を実施** - スライド作成後、必ずPDF変換と画像化で各スライドを視覚的に確認する。特に以下を重点的にチェック：
   - 箇条書きの最後の項目が下部の装飾と重なっていないか
   - テキストがテキストボックスからはみ出していないか
   - 背景や装飾要素とテキストが重なっていないか
   
   ```bash
   # 必ず実行
   libreoffice --headless --convert-to pdf --outdir . working_presentation.pptx
   pdftoppm working_presentation.pdf check -png -r 150
   # 全スライドをviewツールで確認
   ```

3. **テンプレートの項目数を尊重** - 元のテンプレートスライドの項目数を確認し、それを大幅に超えないようにする。超える場合は項目を統合するか、複数スライドに分割

4. **常にテンプレートをコピーして作業** - 元のテンプレートファイルは保持

5. **段階的な作成とチェック** - 一度に全部作らず、2-3スライドごとに視覚的確認を実施

6. **ユーザーフィードバックの早期収集** - レイアウト選択後、作成前に確認

7. **調整の記録** - 行った変更を記録し、最後に報告

---

## 制約事項

- このスキルはPython-pptxライブラリに依存します
- 非常に複雑なテンプレート（アニメーション、埋め込みビデオなど）は完全に保持できない場合があります
- マクロ付きテンプレート（.pptm）は対応していません

---

## 参考資料
- Python-pptx documentation: https://python-pptx.readthedocs.io/