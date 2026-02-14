# Template-Based Presentation Generator

PowerPointのテンプレートスライドに準拠したスライドを自動生成するAgent Skillの開発プロジェクト。

## 概要

このプロジェクトは、PowerPointテンプレートのレイアウトを解析し、そのレイアウトに従って新しいスライドを生成する機能を提供することを目指しています。[python-pptx](https://python-pptx.readthedocs.io/)ライブラリを使用して、テンプレートページをコピーして本文を書き換える方式で実装されます。

## 主な機能（計画中）

- PowerPointテンプレートからスライドレイアウトの解析
- プレースホルダーの識別と抽出
- テンプレートに準拠した新規スライドの生成
- テキスト、画像、表、チャートなどのコンテンツ挿入

## 開発状況

現在、以下の調査・実装が進行中です：

- ✅ PowerPointファイルの基本構造の理解
- ✅ スライドレイアウトとプレースホルダーの解析
- ✅ スライドの削除機能の実装
- 🚧 テンプレートベースのスライド生成機能
- 🚧 Agent Skillとしての統合

## セットアップ

### 前提条件

- Python 3.13以上
- uv（Pythonパッケージマネージャー）

### インストール

```bash
# 依存関係のインストール
uv sync
```

## 依存関係

- **python-pptx** (>=1.0.2) - PowerPointファイルの作成・編集ライブラリ

## リファレンス

`generating-template-based-presentation/references/`ディレクトリには、python-pptxライブラリの使用方法に関する詳細なドキュメントが含まれています：

- **python-pptx-concepts.md** - Presentation、Slide、Layout、Placeholder、Shape、テキスト操作などの基本概念
- **python-pptx-charts-tables.md** - チャートとテーブルの追加・カスタマイズ方法

## 開発者向けメモ

- Pythonの実行には`uv run`を使用
- 依存ライブラリの追加は`uv add`で行う
- パッケージ管理はuvで統一
