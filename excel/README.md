# Gibbs Equilibrium Calculator — Excel (VBA + Solver)

Excel Solver (GRG Nonlinear) を使ってGibbs自由エネルギー最小化で平衡組成を求める
VBAマクロ。Python/Cantera 版と同じNASA7係数・同じ定式化のため結果もほぼ一致する。

## セットアップ

1. **Excel の Solver アドインを有効化**
   - `File → Options → Add-ins → Manage: Excel Add-ins → Go... → Solver Add-in` にチェック

2. **空の Excel ブックを開く**
   - `.xlsm` として新規保存（マクロ有効ブック）

3. **VBA モジュールをインポート**
   - `Alt + F11` で VBE を開く
   - `File → Import File...` で `GibbsEquilibrium.bas` を選択

4. **Solver への参照を追加**（`Application.Run` 経由のため通常は不要だが、
   早期バインディングで動かしたい場合のみ）
   - VBE の `Tools → References...` → `Solver` にチェック

## 使用手順

1. **初回だけ**: マクロ `SetupWorkbook` を実行
   (`Alt + F8 → SetupWorkbook → Run`)
   → `Data`, `Input`, `Solve`, `Output` の4シートが自動生成される

2. **`Input` シート**を編集:
   - B3: 温度 [°C]
   - B4: 圧力 [atm]
   - 各化学種の Inlet moles (B7〜) と Enabled フラグ (C7〜)

3. マクロ `RunEquilibrium` を実行
   → Solver が GRG Nonlinear で解き、`Output` シートに結果を書き込む

## シート構成

| シート | 役割 |
|---|---|
| Data | NASA7係数 (17種、`a1_lo..a7_lo / a1_hi..a7_hi`、元素組成、出典) |
| Input | ユーザー入力 (T, P, inlet, enabled) |
| Solve | 決定変数 `n_i`、目的関数 `G_total/RT`、元素収支 |
| Output | 結果 (converged状態、mol / mol% / G値) |

### 名前付き範囲 (Solver から参照)
- `n_vars` : 決定変数 (各化学種のモル数)
- `G_total` : 最小化目的 (G/RT)
- `nTot_gas` : 気相総モル
- `elem_target`, `elem_actual` : 元素収支の等式制約

## 定式化

最小化: $G/RT = \sum_i n_i \cdot \mu_i/RT$

$$
\mu_i/RT = \begin{cases}
g_i^0/RT + \ln(y_i) + \ln(P/P^\circ) & \text{気相} \\
g_i^0/RT & \text{固体 (activity } \equiv 1)
\end{cases}
$$

$g_i^0/RT$ は NASA7 多項式から:

$$
\frac{g^0}{RT} = a_1(1 - \ln T) - \frac{a_2 T}{2} - \frac{a_3 T^2}{6} - \frac{a_4 T^3}{12} - \frac{a_5 T^4}{20} + \frac{a_6}{T} - a_7
$$

**制約**: $\sum_i a_{ij} n_i = b_j$ (各元素 C/H/O/N の収支)、$n_i \geq 10^{-20}$

**Solver呼び出し**:
```vba
SolverOk  SetCell:=G_total,    MaxMinVal:=2 (Min),  ByChange:=n_vars,  Engine:=1 (GRG)
SolverAdd CellRef:=n_vars,     Relation:=3 (>=),    FormulaText:="1E-20"
SolverAdd CellRef:=elem_actual,Relation:=2 (=),     FormulaText:=elem_target
SolverSolve UserFinish:=True
```

## 収録化学種 (17種)

| 化学種 | 相 | 出典 |
|---|---|---|
| H2, CO2, CO, H2O, CH4, CH3OH, CH3CHO, HCHO, C2H6, C2H4, N2 | 気 | GRI-Mech 3.0 (gri30.yaml) |
| CH3OCH3 | 気 | NASA Glenn L12/92 (nasa_gas.yaml) |
| CH3COOH, C2H5OH, HCOOH | 気 | NASA Glenn L 8/88 (nasa_gas.yaml) |
| CH3COOCH3 | 気 | Burcat T10/07 "Meacetate" |
| C(gr) | 固 | NASA Glenn condensed JANAF X 4/83 |

固体 C(gr) は activity≡1 として扱われる。気相と同じ最小化系に入れると
Solver が `n_{C(gr)} \geq 0` 制約下で適切に 0 or 正値を選ぶ。

## Python/Cantera との比較 (検証値)

H2:CO = 2:1 feed, 250°C, 5 MPa での DME 合成系:

| 種 | Cantera (Python) | Excel Solver (期待値) |
|---|---|---|
| H2 | 0.009 | ~0.009 |
| CO | 0.123 | ~0.123 |
| CO2 | 0.084 | ~0.084 |
| H2O | 0.084 | ~0.084 |
| CH3OH | 0.030 | ~0.030 |
| CH3OCH3 | 0.316 | ~0.316 |

（Solverの収束許容次第で末尾数桁は変動）

## トラブルシューティング

- **Solver が見つからない**: `File → Options → Add-ins` で Solver Add-in が
  有効か確認。VBE の `Tools → References` でも可。
- **収束しない**: `Input` の Enabled を絞る（使わない種は 0 に）。初期値を
  変えて再実行。温度/圧力が極端な条件だと GRG が局所解を返すことがある。
- **`Application.Run "SolverReset"` でエラー**: Solver アドインがブックに
  関連付けされていない。一度 VBE で `Tools → References → Solver` を
  チェックしてから実行すると後続で動くようになる。
- **G_total が `#NUM!`**: enabled 種が1つもない、または n_i 初期値が 0。
  `SetupWorkbook` をもう一度実行してシードし直す。

## 制限

- 非理想気体 (PR/RK EOS) は未対応。理想気体のみ。Python/JS 版と異なる。
- Solver の既定は200変数まで。本マクロは17種なので問題なし。
- 数値精度は Solver Precision 設定 (既定 1e-7) に依存。Python/Cantera の
  1e-10 オーダーよりやや粗い末尾桁差が出ることがある。
