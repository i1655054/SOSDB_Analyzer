# SOSDB Analyzer

SOSDB業務票登録ツール向け VBA解析ツール

## 概要

Excel VBA プロジェクトを解析し、以下の情報を可視化するためのツールです。

- Procedure一覧作成（ProcList）
- 関数トレース
- 関数依存関係解析
- PathChart生成
- VBA解析支援機能

---

## 主な機能

### ProcList

VBAプロジェクト内の Procedure 一覧を生成します。

取得内容

- Module名
- Procedure名
- 種別
- 開始行
- 終了行

### 関数トレース

指定した Procedure の呼び出し先を追跡します。

### 関数依存関係

Procedure 間の依存関係を調査します。

### PathChart

呼び出し経路をツリー形式で可視化します。

対応予定

- frmPathChart
- frmPathChartOption
- 条件指定解析
- ノード数集計
- Edge数集計

---

## UserForms

### frmNavigator

ナビゲーション画面

機能

- シート一覧
- シート検索
- 履歴管理
- 戻る／進む
- モデルレス表示

### frmPathChart

PathChart生成画面

機能

- 開始Procedure指定
- 最大深度指定
- 解析実行
- ログ表示

### frmPathChartOption

PathChart詳細設定

機能

- 解析対象設定
- 除外条件設定
- 出力設定
- 制限設定

---

## プロジェクト構成

```
SOSDB_Analyzer
│
├─ src
│  ├─ Classes
│  ├─ Forms
│  └─ Modules
│
├─ xlsm
│  └─ SOSDB解析ツール.xlsm
│
└─ docs
```

---

## 開発環境

- Microsoft Excel VBA
- Git
- GitHub

---

## 更新履歴

### 2026-08

- GitHub管理開始
- frmNavigator追加
- 履歴管理機能追加
- frmPathChart追加
- frmPathChartOption追加

---

## Author

伊藤正春
