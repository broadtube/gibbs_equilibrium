# Gibbs Equilibrium Calculator — Excel (VBA + Solver)

Excel Solver (GRG Nonlinear) で Gibbs 自由エネルギーを最小化して平衡組成を
求める VBA マクロ。2種類のモジュールを同梱:

| ファイル | EOS | 内容 |
|---|---|---|
| `GibbsEquilibrium.bas` | 理想気体のみ | 直接最小化 1パス |
| `GibbsEquilibrium_PR.bas` | **ideal-gas / Peng-Robinson / Redlich-Kwong** | Picard ループ (phi 固定・再解) |

Python/Cantera と同一の NASA7 係数・定式化で、主要種はほぼ完全一致することを
検証済み (下記「検証」節)。

## 共通セットアップ

1. **Solver アドインを有効化**
   `File → Options → Add-ins → Manage: Excel Add-ins → Go... → Solver Add-in` にチェック

2. **空の `.xlsm` ブック**として新規保存（マクロ有効ブック）

3. **VBA モジュールをインポート**
   - `Alt + F11` で VBE を開く
   - `File → Import File...` → **どちらか一方** を選択
     - `GibbsEquilibrium.bas`  ← 理想気体のみで速く済ませたい場合
     - `GibbsEquilibrium_PR.bas` ← PR/RK 対応、高圧条件で精度が必要な場合
   - （両方インポートしても OK。Sub 名と Module 名が重複しないよう `_PR` サフィックス付き）

4. **Solver 参照** (任意)
   `Application.Run` 遅延バインディング経由なので不要。早期バインディングで
   IntelliSense を効かせたい場合のみ VBE の `Tools → References → Solver`

## 使い方

### 理想気体版 (`GibbsEquilibrium.bas`)
1. `Alt+F8 → SetupWorkbook` 実行 → 4シート (Data / Input / Solve / Output) 生成
2. Input シートを編集: 温度 B3, 圧力 B4, inlet moles B7〜, enabled C7〜
3. `Alt+F8 → RunEquilibrium`

### 非理想気体版 (`GibbsEquilibrium_PR.bas`)
1. `Alt+F8 → SetupWorkbook_PR`
   → 4シート生成 (Data には Tc/Pc/ω 列追加、Solve には phi_i 列追加)
2. Input シートを編集:
   - B3: 温度 [°C]
   - B4: 圧力 [atm]
   - **B5: EOS** (プルダウン: `ideal-gas` / `Peng-Robinson` / `Redlich-Kwong`)
   - B8〜: inlet moles, C8〜: enabled
3. `Alt+F8 → RunEquilibrium_PR`
   - Stage 1: phi=1 で Solver 実行
   - Stage 2+: Solver 結果組成から phi を計算、phi 列に書き込み、再 Solver 実行
   - `|Δφ| < 1e-5` または 10 パスで終了

## シート構成

### 理想気体版 (`GibbsEquilibrium.bas`)
| シート | 列 |
|---|---|
| Data | A=種, B=相, C-E=T範囲, F-L=a_lo, M-S=a_hi, T-W=C/H/O/N, X=Source |
| Input | T, P, inlet, enabled |
| Solve | n_i, g0/RT, μ/RT, 元素収支 |
| Output | mol / mol% / Solver ステータス |

### 非理想版 (`GibbsEquilibrium_PR.bas`)
| シート | 追加/変更列 |
|---|---|
| Data | +X=Tc [K], +Y=Pc [Pa], +Z=omega, AA=Source |
| Input | +B5=EOS selector |
| Solve | +K=phi_i (VBA が更新) |
| Output | +EOS / Picard passes / max \|Δφ\| / Z(gas) / phi_i 列 |

## 定式化

最小化: $G/RT = \sum_i n_i \cdot \mu_i/RT$

理想気体:
$$
\mu_i/RT = \begin{cases}
g_i^0/RT + \ln(y_i) + \ln(P/P^\circ) & \text{気相} \\
g_i^0/RT & \text{固体 (activity ≡ 1)}
\end{cases}
$$

非理想 (PR/RK):
$$
\mu_i/RT = g_i^0/RT + \ln(y_i) + \ln(P/P^\circ) + \ln(\varphi_i) \quad \text{気相}
$$

$g_i^0/RT$ は NASA7 多項式から:
$$
\frac{g^0}{RT} = a_1(1 - \ln T) - \frac{a_2 T}{2} - \frac{a_3 T^2}{6} - \frac{a_4 T^3}{12} - \frac{a_5 T^4}{20} + \frac{a_6}{T} - a_7
$$

**制約**: $\sum_i a_{ij} n_i = b_j$ (各元素 C/H/O/N の収支)、$n_i \geq 10^{-20}$

### PR/RK フガシティ係数

混合則 ($k_{ij}=0$): $a_\text{mix} = \left(\sum_i y_i \sqrt{a_i}\right)^2$, $b_\text{mix} = \sum_i y_i b_i$

PR per-species:
- $\kappa_i = 0.37464 + 1.54226\omega_i - 0.26992\omega_i^2$
- $\alpha_i(T) = [1 + \kappa_i(1 - \sqrt{T/T_{c,i}})]^2$
- $a_i = 0.45724 (R T_{c,i})^2 / P_{c,i} \cdot \alpha_i$
- $b_i = 0.07780 R T_{c,i} / P_{c,i}$

RK per-species:
- $a_i = 0.42748 R^2 T_{c,i}^{2.5} / P_{c,i}$
- $b_i = 0.08664 R T_{c,i} / P_{c,i}$

3次方程式を Cardano で解いて気相根 $Z$ を取得、$\ln\varphi_i$ を以下で:

PR:
$$
\ln\varphi_i = \frac{b_i}{b_\text{mix}}(Z-1) - \ln(Z-B) - \frac{A}{2\sqrt{2}B}\left(\frac{2\sqrt{a_i}\sum_j y_j \sqrt{a_j}}{a_\text{mix}} - \frac{b_i}{b_\text{mix}}\right)\ln\frac{Z+(1+\sqrt{2})B}{Z+(1-\sqrt{2})B}
$$

RK: 同形式で、最後の $\ln((Z+(1+\sqrt2)B)/(Z+(1-\sqrt2)B))$ を $\ln(1 + B/Z)$ に置換、$A/(2\sqrt2 B) \to A/B$。

### なぜ Picard ループ (固定 φ) で UDF 呼出しを避けるか
- UDF を μ/RT 式に直接埋め込むと、GRG iteration ごとに cubic solve が走り **遅い + 数値不安定**
- PR/RK の vapor/liquid root 切替点で GRG の数値勾配が破綻する
- 対して **Picard**: Stage 毎に phi を固定 → 目的関数は滑らかな対数線形 → GRG が得意な問題形式
- 典型 2-3 パスで $|\Delta\varphi| < 10^{-5}$ に収束 (検証済)

## 収録化学種 (17種)

| 化学種 | 相 | NASA7 出典 | Tc/Pc/ω 出典 |
|---|---|---|---|
| H2, CO, CO2, H2O, CH4, CH3OH, CH3CHO, HCHO, C2H6, C2H4, N2 | 気 | GRI-Mech 3.0 (gri30.yaml) | NIST WebBook / IAPWS-95 (H2O) / Reid-Prausnitz-Poling |
| CH3OCH3 | 気 | NASA Glenn L12/92 | NIST |
| CH3COOH, C2H5OH, HCOOH | 気 | NASA Glenn L 8/88 | NIST |
| CH3COOCH3 | 気 | **Burcat T10/07 "Meacetate"** | chemicals lib (DIPPR) |
| C(gr) | 固 | NASA Glenn JANAF X 4/83 | (N/A, activity=1) |

## Python/Cantera との比較 (検証)

VBA の PR/RK 数理を Python に移植、`scipy.optimize.minimize` + Picard ループで解いた結果を Cantera PR/RK 平衡と比較:

| 条件 | EOS | Picard passes | 主要種 max rel | phi レンジ |
|---|---|---|---|---|
| H2:CO 2:1, 250°C, 5 MPa | PR | 3 | 2.6×10⁻⁷ | 0.87–1.06 |
| H2:CO 2:1, 200°C, 10 MPa | PR | 3 | 5.3×10⁻⁶ | 0.67–1.18 |
| H2:CO 2:1, 250°C, 1 atm | PR | 2 | 7×10⁻⁷ | ~1.00 |
| H2:CO 2:1, 250°C, 5 MPa | RK | 3 | 2.8×10⁻⁷ | 0.88–1.08 |
| H2:CO 2:1, 200°C, 10 MPa | RK | 3 | 4.5×10⁻⁶ | 0.70–1.24 |

→ 主要種は **6-7 桁一致**。微量種 (mol < 10⁻⁴) で 2-7% 差があるが Solver 許容差 (ftol=1e-10) 由来でありアルゴリズム誤差ではない。

（理想気体版は従来通り Cantera と 10⁻⁶ 相対一致。Picard 前の Stage 1 結果と同じ）

## トラブルシューティング

- **Solver が見つからない**: `File → Options → Add-ins` で Solver Add-in を有効化
- **Picard が収束しない** (`|Δφ|` が tolerance 以上で 10 passes 消費): 条件が相境界近傍。`EOS: ideal-gas` で近似するか、Python/Cantera に移行
- **`Application.Run "SolverReset"` でエラー**: Solver アドインがブックに関連付けされていない → VBE で一度 `Tools → References → Solver` をチェックしてから実行
- **`#NUM!` が μ/RT セルに**: enabled 種が 0、または n_i の初期値が 0 → `SetupWorkbook_PR` 再実行で初期値リセット
- **Z < B エラー**: 液相根が選ばれた可能性。温度を上げて気相域に戻すか、ideal-gas にフォールバック

## 制限

- **液相/二相分離は扱えない**: 全て気相仮定 (graphite のみ固体として別扱い)
- **k_ij = 0** 固定: 極性分子 (H2O, CH3OH) 混在で精度低下の可能性。値を調整したい場合は `ComputeFugacities` 内の mixing 式を修正
- Solver 既定 Precision=1e-7 → 末尾数桁は Cantera (1e-10) より粗い
- 200変数上限の Excel Solver だが、本マクロは 17 種なので余裕あり
- **SRK (Soave-Redlich-Kwong), BWR, virial は非対応** (Cantera も native 非対応)
