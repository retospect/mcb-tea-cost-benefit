# MCB TEA and Cost Benefit

Marine Cloud Brightening: Cost, Benefit, and Planetary-Scale Assessment.

A bottom-up techno-economic assessment (TEA) of marine cloud brightening in a single Excel workbook. It uses the best available academic sources for inputs and connects them into one end-to-end assessment, then works through different delivery mechanisms: a dedicated vessel, piggybacking on a cargo ship, fixed platforms, and stationary vessels. The foundation is the numbers from Daniel Harrison's prototype on the Great Barrier Reef.

**Workbook:** `Draft MCB TEA and Cost Benefit TJ 10122026.xlsx` (draft: global benefit numbers may change once the eligible-particle cutoff is settled)

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/titojankowski/mcb-tea-cost-benefit/blob/main/colab/mcb_tea.ipynb)

## How the workbook is organized

Start on the `Overview` tab. The workbook has two parts.

**Particle cost model (tabs 1-4): what it costs to produce and spray particles**

| Tab | What it covers |
| --- | --- |
| 1. Technology Description | The spray machine. CCN product = 30-1000 nm particles (Harrison 2025 convention). |
| 2. Cost Inputs | The blue input cells: capital, energy, crew, unit costs. |
| 3a. Outputs - Dedicated Vessel | Levelized cost per tonne of aerosol (the anchor number). |
| 3b / 3c / 3d | Cheaper siting and delivery vs a dedicated vessel: ship piggyback, fixed platform, stationary vessel. |
| 4. Scenario Comparison | Levelized cost per tonne of aerosol across 3a-3d, side by side. |

**Cooling costs and benefit (tabs 5-6): what the cooling is worth**

| Tab | What it covers |
| --- | --- |
| 5a / 5b | $/tonne into $/unit cooling, the model the benefit tabs read (5a); sensitivity tornado of that answer (5b). |
| 5c / 5c-alt / 5d | Local benefits: reef bleaching relief (5c, plus a Zhao-based alternate) and city heat deaths (5d). |
| 6a | Tonnes into global cooling via published Earth-system models. |
| 6b / 6c | Value per degree-year vs climate damage functions, for the dedicated vessel and piggyback cases. |
| References | Every paper the inputs were pulled from. |

## Legend

- Black text = formula
- Blue fill = inputs, you can change these
- Green text = pulled from another tab

## Out of scope

- Monitoring, reporting, verification (MRV): no measured forcing, no verification, no certification.
- Marine ecosystem impacts other than cooling and its effects.
- Large-scale circulation response, i.e. what cooling the Great Barrier Reef by 1C does to other parts of the climate system.

## Questions, comments, complaints

Tito Jankowski, tito.jankowski@renphil.org
