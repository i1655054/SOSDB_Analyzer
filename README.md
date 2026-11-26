# SOSDB Analyzer

Current Version : v2.0.0

---
## 概要

SOSDB問診票登録ツールのVBAソースコードを解析し、 プロシージャ一覧・呼出関係・PathChart・HCPチャートを生成する解析ツールです。

---

# 主な機能

- Procedure一覧生成
- Function Trace生成
- Function Dependency生成
- PathChart生成
- PathChartTree生成
- PowerPoint出力
- PNG出力
- HCPチャート生成
- HCP_TARGETS管理
- HCP一括生成
- hcp.txt出力
- hcp.html出力

## Procedure一覧生成

VBAプロジェクトを解析し、Procedure情報を収集します。

出力項目

- Module名
- Procedure名
- Procedure種別
- JPName
- Category
- Input
- Output
- Summary
- Remarks
- UsageCount

生成シート

- ProcList

---

## Function Trace生成

Procedureの呼出関係を解析し、トレース情報を生成します。

生成シート

- 関数トレース

---

## Function Dependency生成

呼出元・呼出先の依存関係を一覧化します。

生成シート

- 関数依存関係

---

## PathChart生成

指定Procedureを起点として呼出ツリーを生成します。

機能

- 呼出階層表示
- モジュール表示
- カテゴリ表示
- 日本語名称表示
- 共通関数除外
- 外部モジュール除外
- 再帰呼出制御

生成シート

- PathChart

---

## PathChartTree生成

PathChartをテキストツリー形式で出力します。

生成シート

- PathChartTree

例

```text
MainProcess
├─ LoadData
│  └─ ReadCsv
├─ Validate
└─ SaveData
```

---

## PowerPoint出力

PathChartをPowerPointファイルとして出力します。

出力形式

- PowerPoint(.pptx)
- PNG(.png)

対応

- 縦向き
- 横向き
- 自動判定

---

## HCPチャート生成

Procedure内部を解析し、 HCP(Hierarchical Control Process)形式の 業務仕様書を生成します。

生成物

- hcp.txt
- hcp.html
- HCPChartシート

対応機能

- Calls一覧生成
- Called By一覧生成
- Procedureリンク生成
- Business説明出力
- HCP_TARGETS管理
- HCP一括生成

対応ノード

- START
- END
- IF
- ELSE
- FOR
- DO LOOP
- SELECT CASE
- WITH
- CALL
- REFERENCE
- NORMAL

生成シート

- ProcNodes
- HCPNodes
- HCPChart

---

# システム構成

```text
src
├─ Analyzer
├─ Repository
├─ PathChart
├─ HCP
├─ PowerPoint
├─ ToolManager
└─ Model

Class
├─ clsAnalyzeContext
├─ clsProcInfo
├─ clsProcNode
├─ clsTreeNode
├─ clsHCPNode
├─ clsToolInfo
├─ clsTextRenderer
├─ clsHtmlRenderer
├─ clsExcelRenderer
└─ clsHCPRenderBuilder
```

# 処理フロー
```text
解析対象WorkbookからProcedure情報を収集し、 依存関係解析、PathChart生成、HCPチャート生成へ展開する。

解析対象Workbook
       ↓
  ProcAnalyzer
       ↓
    ProcList
       ↓
 FunctionTrace
       ↓
FunctionDependency
       ↓
   PathChart
       ↓
 PathChartTree


解析対象Workbook
       ↓
  ProcAnalyzer
       ↓
   ProcNodes
       ↓
   HCPNodes
       ↓
   HCPChart


PathChart
    ↓
PowerPoint出力
    ↓
PPTX / PNG
```
---

# 管理ブック構成

```text
ProcList
関数トレース
関数依存関係
NoCallerProcList

PathChart
PathChartTree

ProcNodes
HCPNodes
HCPChart
HCP_TARGETS (HCP一括生成管理)
```
---
## HCP_TARGETS管理

HCP生成対象Procedureを管理します。

機能

- ProcListとの同期
- Procedure追加
- Procedure削除同期
- HCP生成対象管理
- HCP一括生成管理

管理項目

- Exec
- ModuleName
- ProcName
- Status
- LastUpdate
- ErrorNo
- ErrorMessage

整合性確認

- CheckHCPTargetIntegrity
---
## HCP生成手順

単体生成

解析実行 ↓ Procedure選択 ↓ HCP生成

一括生成

解析実行
 ↓ 
HCP_TARGETS更新
 ↓ 
Exec列へ Y 設定
 ↓ 
HCP一括生成

生成結果は HCP_TARGETSへ記録される

- Status
- LastUpdate
- ErrorNo
- ErrorMessage

---

# アーキテクチャ

## AnalyzeContext

実行中の解析情報を管理します。

- TargetBook
- ManagementBook
- ToolInfo
- ProcedureRepository

## HCPアーキテクチャ

```text
Procedure
 ↓ 
AnalyzeProcedure
 ↓ 
ProcNodes
 ↓ 
ConvertProcToHCP
 ↓ 
HCPNodes
 ↓ 
clsHCPRenderBuilder
 ↓ 
RenderText
 ↓ 
Renderer

├─ clsTextRenderer
├─ clsHtmlRenderer
└─ clsExcelRenderer

出力

├─ hcp.txt
├─ hcp.html
└─ HCPChart
```
---

## ProcedureRepository

ProcListシートからProcedure情報をキャッシュし、高速検索を提供します。

機能

- Procedure情報取得
- Procedure解決
- Procedure索引管理

---

# 開発方針

- Option Explicit必須
- Public ProcedureはJPNameコメント必須
- Summaryコメント必須
- 変数コメント記述推奨
- モジュール責務分離
- Repositoryパターン採用
- Contextパターン採用

---

# 動作環境

- Microsoft Excel VBA
- Windows
- PowerPoint（PowerPoint出力時）

---

# リリース履歴

## v1.0.0

- ProcedureRepository導入
- AnalyzeContext導入
- PathChart機能整理
- PowerPoint出力対応
- PNG出力対応
- モデルクラス整理
- コメント整備
- 保守性改善

## v2.0.0

HCP生成機能の全面リファクタリング

Renderer分離により Text / HTML / Excel出力をクラス化し 保守性および拡張性を改善

主な変更点

- clsTextRenderer追加
- clsHtmlRenderer追加
- clsExcelRenderer追加
- clsHCPRenderBuilder追加
- modHCPExporter廃止
- HCP_TARGETS同期改善
- CheckHCPTargetIntegrity追加
- HCP一括生成安定化
- 回帰テスト整備


---

# 作成者

SOSDB Analyzer Project
