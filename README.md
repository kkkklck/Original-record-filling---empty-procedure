# Original Record Auto-Filling Suite

**Proprietary commercial software for structural fireproofing inspections.**

The Original Record Auto-Filling Suite streamlines how inspection agencies compile fireproofing records for steel columns, beams, braces, and space frames. Starting from Word-based raw reports, the suite extracts measurements, generates a consolidated Word summary, and writes normalized Excel deliverables that are ready for review and printing. The underlying automation script (`原始记录自动填写程序.py`) is offered only under commercial license.

> **Closed-source / license required.** Copyright registered. Source code is not distributed.

## Highlights
- **Word-to-structure extraction** – Detects tables that include entries such as “测点1 / 平均值”, tolerates merged or missing columns, and outputs a highlighted summary document with consistent row counts.
- **Excel batch publishing** – Clones template sheets while preserving print areas, margins, zoom, frozen panes, and cell sizing, then writes data in-place with layout fidelity.
- **Smart categorization & sorting** – Automatically classifies by steel column, beam, brace, and space frame (with synonyms). Floors are ordered `B* → numerical → equipment → roof`.
- **μ-value routing** – Flags a row as μ when any reading contains four or more consecutive digits or an absolute value ≥ 1000. Normal and μ pages share a continuous numbering scheme without mixing pages.
- **Dynamic page pooling** – Copies template sheets only when necessary, keeps numbering aligned between normal and μ pages, and removes unused μ templates.
- **Typography enforcement** – Ensures cells containing “μ” use Times New Roman while keeping all other styles intact.
- **Operator-friendly experience** – Guided workflow with progress feedback, safe recovery when files are locked, and clear prompts throughout the process.

## Built-in Operating Modes
1. **Date buckets (default)** – Distributes components across dates with a choice of front-loaded or back-loaded prioritization.
2. **Floor breakpoints** – Map custom floor ranges (e.g., 3/6/equipment/roof) to specific dates.
3. **Single-day** – Produce all deliverables under one inspection date for rapid submission.
4. **Floor × date slicing** – Share a floor across multiple dates evenly or by quota.

Braces and space frames can be grouped by identifier or floor; space-frame subtypes support whitelist constraints.

## Applicable Scenarios
- Bulk record keeping and report generation for structural fireproofing or similar inspections covering steel columns, beams, braces, and space frames.
- Agencies that require one-click consolidation with regulation-compliant formatting and print-ready output.

## System Requirements
- Windows 10 or 11 workstations for operation and printing.
- Python 3.9+ runtime with `python-docx` and `openpyxl` dependencies (included in demo builds).
- Excel templates that contain base sheets (e.g., "钢柱", "钢梁", "支撑", "网架") and the corresponding μ variants (e.g., "钢梁μ").

## Engagement & Evaluation
This repository does not include source code. To request a demo or trial:
1. Share your target structures, template samples, and required layout.
2. Receive a remote walkthrough or sanitized demo package.
3. Finalize scope and SLA; sign the commercial agreement.
4. Receive the protected executable/script bundle and aligned templates.
5. Complete training, acceptance, and transition into maintenance.

## Licensing & Services
- **License models** – Per-site license, annual maintenance, or one-time buyout (includes template assets, secondary-development hooks, and priority support).
- **Deliverables** – Executable (or controlled script), templates and samples, user manual, and go-live assistance.
- **Professional services** – Template adaptation, field mapping, process adjustments, custom rules (ID/floor/date logic), and on-premise deployment.
- **Pricing** – Quoted based on feature scope and delivery schedule; milestone-based payment supported.

> The software is protected under copyright law. Redistribution, reverse engineering, or sublicense without written consent is prohibited. Interface extensions can be negotiated contractually.

## FAQ (Condensed)
- **Can additional categories or fields be supported?** Yes, through customized mapping.
- **Can the μ-detection logic be adjusted?** Thresholds, patterns, and units are configurable per project.
- **Is a fixed template required?** Layouts are adaptable, but a final print template must be provided.
- **Does the workflow require Microsoft Office?** Generation does not rely on Word/Excel processes, but Office is recommended for review and printing.

## Roadmap
- Custom field validation with anomaly highlighting.
- Finer-grained hooks for floor and identifier sorting.
- Graphical wizard with audit logging.
- Bilingual (Chinese/English) interface options.

## Release Notes
- **v1.0.1** – Stable release featuring enhanced pagination and μ-routing, template cloning with preserved print/zoom settings, typography fidelity, and guided progress indicators.

## Contact
- **Email** – <kkkklck@qq.com>
- **WeChat** – `wunailck`
- Or open an issue describing your organization and intended use (source code will not be provided; issues are handled for business communication only).

## Legal Notice
This repository is solely for product information and business coordination. Any reverse engineering, competitive training, or replication without written authorization is forbidden and will trigger legal action.
