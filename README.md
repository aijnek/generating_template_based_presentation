# Template-Based Presentation Generator

PowerPointテンプレートに準拠したスライドを自動生成するAgent Skillプロジェクト。

## 概要

テンプレートPPTXファイルのレイアウトを解析し、[python-pptx](https://python-pptx.readthedocs.io/)を使ってテンプレートデザインに準拠した新規スライドを生成します。

## セットアップ

```bash
uv sync
```

## 使い方

スキルとして使用する際は `generating-template-based-presentation/SKILL.md` を参照してください。

## プロジェクト構成

- `generating-template-based-presentation/SKILL.md` - Agent Skill定義とワークフロー
- `src/analyze_ppt.py` - PPTXファイル解析用のユーティリティスクリプト
- `pyproject.toml` - プロジェクト設定と依存関係
