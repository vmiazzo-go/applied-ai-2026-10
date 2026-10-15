# NovaHome `product_analytics` — Presentazione Didattica (38 Slide)

Presentazione didattica (38 slide, formato 16:9 widescreen, 100% in italiano, stile **Official Google Cloudstyle 2.0**) progettata per accompagnare la spiegazione delle query avanzate BigQuery AI del laboratorio **Lab II: Analisi Intelligente delle Recensioni (`product_analytics`)**.

- **Presentazione Google Slides**: [`1h0FrF2z0yW_c-se6FBIXlIkdKt3Xkns58QsZDuJVmsg`](https://docs.google.com/presentation/d/1h0FrF2z0yW_c-se6FBIXlIkdKt3Xkns58QsZDuJVmsg/edit)

## Regole di Design Rispettate
- **4 Pagine per ogni Modulo Query** (senza numerazione ridondante `1/4` nei badge o nei footer):
  1. **Obiettivo di Business** (Problema, Domanda Chiave, Impatto sul caso NovaHome)
  2. **Diagramma E/R (`Mermaid erDiagram`) + Post-it AI**:
     - Le relazioni tra tabelle usano la notazione Crow's Foot (`||--||`, `||--o{`, `}o--o{`) riportando **esclusivamente molteplicità e chiave** (`1:1 review_id`, `1:N sku`, `1:N batch_id`, `1:N lot_id`, `1:N machine_id`, `1:N component_type`).
     - Le proiezioni in sola lettura (`SELECT` / viste / TVF) sono modellate con una **freccia direzionale semplice `───► SELECT`** (senza zampe di gallina E/R).
     - I dettagli sulle funzioni AI invocate e sulle trasformazioni sono raccolti in un **Post-it giallo stilizzato** a destra del diagramma.
  3. **Particolarità Tecniche** (Pattern architetturali, ottimizzazione costi, gestione quote)
  4. **Listato SQL Sintetico** ($\le 18$ righe in `Roboto Mono` $\ge 10\text{pt}$ con istruzioni chiave evidenziate e callout laterali)
- **Speaker Notes complete in italiano** su tutte le 38 slide con approfondimenti didattici (perché non usare sempre `AI.CLASSIFY`, perché `AI.GENERATE_INT` non supporta `MINIMIZE_COST` e richiede `endpoint`, differenza tra `AI.SCORE` e `AI.GENERATE_INT` in `ORDER BY`, e chiave primaria composta su `batch_components`).
- **Zero muri di testo** ("meno è meglio") e **font minimo $\ge 10\text{pt}$** verificato su tutte le 38 pagine.
- **Codifica Semantica a 4 Colori**:
  - 🔵 **Blu (`#1967D2`)**: Dati grezzi (`product_reviews`) & Funzioni BigQuery AI (`AI.IF`, `AI.CLASSIFY`, `AI.SCORE`, `AI.GENERATE`, `AI.KEY_DRIVERS`, `AI.AGG`, `AI.SIMILARITY`, `AI.FORECAST`)
  - 🟢 **Verde (`#137333`)**: Tabelle arricchite dall'AI (`review_insights`, viste) & Schema tipizzato / parametri (`output_schema`, `MINIMIZE_COST`, `TimesFM 2.5`)
  - 🟡 **Giallo / Ambra (`#B06000`)**: Grafo Supply Chain / Distinta Base (`products`, `production_batches`, `batch_components`, `component_lots`) & Join relazionali (`LEFT JOIN ... IS NULL`, `GROUP BY`)
  - 🔴 **Rosso (`#C5221F`)**: Causa radice fisica isolata (`machines` $\rightarrow$ `CAL-02`, lotto `HE-4471` al 77%) & Resilienza quote (`REPEAT ... UNTIL`, `retryable error`)

## File Inclusi nella Cartella (Versionabili in Git)
- [`er_diagrams.mmd`](./er_diagrams.mmd): Sorgenti Mermaid `erDiagram` per tutti i 10 diagrammi E/R del deck (Mappa Completa + 9 Moduli).
- [`er_img/`](./er_img/): I 10 diagrammi Mermaid `erDiagram` pre-renderizzati sia in formato vettoriale (`.svg`) sia ad alta risoluzione 2x (`.png`) con codifica semantica a 4 colori e frecce `SELECT`.
- [`slides.html`](./slides.html): Presentazione standalone in HTML5/CSS3 con i diagrammi Mermaid `erDiagram` SVG e i Post-it incorporati (100% testuale e diffabile in Git; apribile nel browser e stampabile in PDF).
- [`product_analytics_deck.pptx`](./product_analytics_deck.pptx): Presentazione PowerPoint nativa esportata da Google Slides con i 10 diagrammi Mermaid `erDiagram`, i Post-it, le Speaker Notes e i blocchi SQL a colori semantici.
- [`product_analytics_deck.pdf`](./product_analytics_deck.pdf): Versione PDF vettoriale pronta per la proiezione in aula.
- [`build_deck.py`](./build_deck.py), [`er_rendered.json`](./er_rendered.json) & [`gslides_batch.json`](./gslides_batch.json): Script generatore in Python e payload delle 38 slide.

## Indice dei 9 Moduli (38 Slide)
- **Slide 1–2**: Copertina + **Mappa E/R Completa** (`KTL-100` $\rightarrow$ `HE-4471` $\rightarrow$ `CAL-02`)
- **Modulo 01 (Slide 3–6)** — `sql/02_quick_triage.sql` (`2a`): *Triage Ottimizzato per i Costi* (`AI.IF` + `AI.CLASSIFY` + `AI.GENERATE_INT`)
- **Modulo 02 (Slide 7–10)** — `sql/02_quick_triage.sql` (`2b`): *Classifica Semantica del Rischio Fisico* (`AI.SCORE` in `ORDER BY`)
- **Modulo 03 (Slide 11–14)** — `sql/03_async_enrichment_pipeline.sql` (`3a`): *Pipeline di Arricchimento Incrementale* (`AI.GENERATE` con `output_schema` a 6 campi, anti-join incrementale e retry loop `429`)
- **Modulo 04 (Slide 15–18)** — `sql/04_root_cause_analysis.sql` (`4a+4b+4c`): *Analisi della Causa Radice sulla Supply Chain* (Join AI $\leftrightarrow$ BOM `affected_component = component_type`, isolamento `HE-4471` al 77% e macchina `CAL-02`)
- **Modulo 05 (Slide 19–22)** — `sql/05_executive_summary.sql` (`5`): *Piano d'Azione Esecutivo da SQL a Prompt* (Pre-aggregazione deterministica `STRING_AGG` + `AI.GENERATE`)
- **Modulo 06 (Slide 23–26)** — `sql/07_key_drivers_and_ai_agg_enhancements.sql` (`7a`): *Scoperta Automatica delle Cause Radice* (`AI.KEY_DRIVERS`, Lift `+1600%`)
- **Modulo 07 (Slide 27–30)** — `sql/07_key_drivers_and_ai_agg_enhancements.sql` (`7b`): *Aggregazione Semantica Multi-Riga Nativa* (`AI.AGG` dentro `GROUP BY sku`)
- **Modulo 08 (Slide 31–34)** — `sql/07_key_drivers_and_ai_agg_enhancements.sql` (`7c+7d`): *Correlazione Semantica Senza Chiavi Esterne* (`JOIN ON AI.IF` vs `AI.SIMILARITY` con `text-embedding-005`)
- **Modulo 09 (Slide 35–38)** — `sql/07_key_drivers_and_ai_agg_enhancements.sql` (`7e+7f+7g`): *Analisi Zero-Shot delle Serie Storiche* (`TimesFM 2.5`: `AI.DETECT_ANOMALIES`, `AI.FORECAST` multivariato e `AI.EVALUATE`)
