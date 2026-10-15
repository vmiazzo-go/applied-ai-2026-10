#!/usr/bin/env python3
"""
Builds the NovaHome Product Review Intelligence (product_analytics) 38-slide Professor Deck.
Outputs:
  1. product_analytics/deck/slides.html       (Standalone 16:9 HTML/CSS/SVG Cloudstyle 2.0 deck, 100% git-diffable)
  2. product_analytics/deck/gslides_batch.json (Deterministic Google Slides API batch payload for PPTX/PDF export)
"""

import html
import json
import os
import re

DECK_DIR = os.path.dirname(os.path.abspath(__file__))

# Official Google Cloudstyle 2.0 & Semantic Palette
COLORS = {
    "white": "#FFFFFF",
    "dark": "#202124",
    "grey": "#5F6368",
    "border": "#DADCE0",
    "light_bg": "#F8F9FA",
    # Semantic 4-color system
    "blue": {
        "accent": "#1967D2",
        "primary": "#4285F4",
        "fill": "#E8F0FE",
        "border": "#8AB4F8",
        "label": "DATI GREZZI & BIGQUERY AI",
    },
    "green": {
        "accent": "#137333",
        "primary": "#34A853",
        "fill": "#E6F4EA",
        "border": "#81C995",
        "label": "INSIGHT AI & SCHEMA TIPIZZATO",
    },
    "yellow": {
        "accent": "#B06000",
        "primary": "#FBBC04",
        "fill": "#FEF7E0",
        "border": "#FDD663",
        "label": "DISTINTA BASE (BOM) & JOIN SQL",
    },
    "red": {
        "accent": "#C5221F",
        "primary": "#EA4335",
        "fill": "#FCE8E6",
        "border": "#F28B82",
        "label": "CAUSA RADICE & RESILIENZA",
    },
}

# ==============================================================================
# 9 CORE QUERY MODULES SPECIFICATION (4 SLIDES EACH = 36 SLIDES + 2 INTRO = 38)
# ==============================================================================
MODULES = [
    {
        "id": "M1",
        "num": "01",
        "script": "sql/02_quick_triage.sql (Query 2a)",
        "name": "Triage Ottimizzato per i Costi",
        "subtitle": "AI.IF + AI.CLASSIFY + AI.GENERATE_INT",
        "goal": {
            "headline": "Separare i veri difetti fisici dal rumore delle recensioni contenendo i costi operativi",
            "cards": [
                {
                    "color": "blue",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "Rumore nelle Recensioni",
                    "bullets": [
                        "Il 65% delle recensioni riguarda spedizioni, gusti personali o giudizi a 5 stelle",
                        "Leggere e classificare manualmente ogni singolo testo non scala su migliaia di reclami",
                        "Il team Qualità ha bisogno di una prima scrematura rapida ed economica",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Tre Segnali per Ogni Reclamo",
                    "bullets": [
                        "La recensione segnala un vero difetto fisico/sicurezza o un reclamo commerciale?",
                        "In quale categoria di guasto rientra il problema (termico, elettrico, meccanico, estetico)?",
                        "Qual è il livello di gravità percepito dal cliente su una scala da 1 a 5?",
                    ],
                },
                {
                    "color": "green",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Triage Rapido e Sostenibile",
                    "bullets": [
                        "Isolare immediatamente i reclami reali di guasto o sicurezza dal resto del catalogo",
                        "Classificare e pesare ogni segnalazione mantenendo minimo il costo unitario di analisi",
                        "Fornire al team Qualità una vista esplorativa immediata sui prodotti più critici",
                    ],
                },
            ],
            "kpi": "Separare Difetti Reali e Rumore  •  Tassonomia a 5 Categorie  •  Stima Urgenza 1–5 a Costo Minimo",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 01 (OBIETTIVO DI BUSINESS):\n\n"
                "1. CONTESTO DI BUSINESS:\n"
                "   • La maggior parte delle recensioni e-commerce non riguarda difetti di fabbricazione ma ritardi del corriere, imballaggi o giudizi positivi.\n"
                "   • Prima di avviare indagini ingegneristiche costose, il team Qualità deve rispondere a 3 domande: è un difetto reale? Di che categoria? Quanto è urgente?"
            ),
        },
        "er": {
            "headline": "Scansione su singola tabella di recensioni grezze con arricchimento AI al volo",
            "postit": {
                "title": "📌 Arricchimento On-the-Fly",
                "lines": [
                    "• Freccia SELECT: proiezione in sola",
                    "  lettura (nessuna tabella scritta)",
                    "• Input: product_reviews.review_text",
                    "• AI.IF(...) ➔ is_defect (BOOL)",
                    "• AI.CLASSIFY(...) ➔ category (STRING)",
                    "• AI.GENERATE_INT(...) ➔ urgency (INT64)",
                ],
            },
            "note": "",
            "notes": "", 
        },
        "tech": {
            "headline": "Funzioni AI specializzate e distillazione automatica del modello riducono il TCO",
            "cards": [
                {
                    "color": "blue",
                    "tag": "FUNZIONI AI SPECIALIZZATE",
                    "title": "Tipi SQL Forti per Ogni Task",
                    "bullets": [
                        "AI.IF restituisce direttamente un BOOL nativo (niente parsing di 'yes'/'true')",
                        "AI.CLASSIFY vincola l'output esattamente alle 5 stringhe dell'array categories",
                        "AI.GENERATE_INT restituisce una STRUCT con .result tipizzato INT64",
                    ],
                },
                {
                    "color": "green",
                    "tag": "OTTIMIZZAZIONE DEI COSTI",
                    "title": "optimization_mode = MINIMIZE_COST",
                    "bullets": [
                        "Supportata sia da AI.IF sia da AI.CLASSIFY per carichi ad alto volume",
                        "Usa un modello teacher Gemini su un campione e distilla un modello student leggero",
                        "AI.GENERATE_INT non usa distillazione: invoca direttamente il modello Flash specificato",
                    ],
                },
                {
                    "color": "red",
                    "tag": "TRADE-OFF ARCHITETTURALE",
                    "title": "Esplorazione vs Produzione",
                    "bullets": [
                        "Nel Modulo 03 vedremo come comprimerle in 1 sola chiamata con AI.GENERATE",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 01 (PARTICOLARITÀ TECNICHE):\n\n"
                "1. PERCHÉ NON USARE SEMPRE AI.CLASSIFY PER TUTTI E 3 I SEGNALI?\n"
                "   • Tipo di ritorno SQL nativo (BOOL vs STRING vs INT64): AI.CLASSIFY restituisce sempre e solo una STRING. "
                "Se usassimo AI.CLASSIFY per il filtro difetti o per l'urgenza, otterremmo stringhe testuali che richiedono CAST espliciti. Invece:\n"
                "     - AI.IF restituisce direttamente un BOOL SQL nativo, immediatamente utilizzabile in WHERE, COUNTIF() o JOIN ON.\n"
                "     - AI.CLASSIFY restituisce una STRING vincolata a una tassonomia chiusa nominale (5 categorie mutuamente esclusive).\n"
                "     - AI.GENERATE_INT restituisce un INT64 nativo (tramite .result), pronto per operazioni matematiche (AVG, >= 4, SUM).\n"
                "   • Calibrazione semantica del task: ogni funzione istruisce il modello con un prompt di sistema ottimizzato per quel compito.\n\n"
                "2. PERCHÉ IN AI.GENERATE_INT INDICHIAMO endpoint => 'gemini-2.5-flash' MENTRE IN AI.IF E AI.CLASSIFY NO?\n"
                "   • AI.IF, AI.CLASSIFY e AI.SCORE sono funzioni 'managed' (task-specific): BigQuery seleziona e gestisce automaticamente il modello ottimale "
                "sotto il cofano (e quando optimization_mode => 'MINIMIZE_COST' è attivo, orchestra sia il modello teacher sia lo student distillato), "
                "pertanto non richiedono il parametro endpoint.\n"
                "   • AI.GENERATE_INT (come tutta la famiglia AI.GENERATE*) offre invece controllo diretto sul modello generativo: "
                "anche se endpoint ha un default se omesso, in produzione è best practice esplicitare endpoint => 'gemini-2.5-flash' per fissare versione, costi e latenza.\n\n"
                "3. PERCHÉ AI.GENERATE_INT NON HA optimization_mode => 'MINIMIZE_COST'? LA DISTILLAZIONE È IMPLICITA?\n"
                "   • No: il parametro optimization_mode => 'MINIMIZE_COST' (distillazione automatica teacher-student) è supportato esclusivamente "
                "dalle funzioni gestite di filtraggio e classificazione AI.IF e AI.CLASSIFY.\n"
                "   • AI.GENERATE_INT non supporta il parametro optimization_mode né esegue distillazione implicita: invoca direttamente il modello Gemini per ogni riga. "
                "Per contenere i costi con AI.GENERATE_INT si utilizza un modello veloce ed economico come 'gemini-2.5-flash'."
            ),
        },
        "sql": {
            "headline": "Query 2a — Triage booleano, classificazione e urgenza intera in un'unica SELECT",
            "code": [
                "SELECT r.review_id, r.sku, r.rating,",
                "  AI.IF(('Defect or safety issue: ', r.review_text),",
                "    connection_id => 'us.vertex-ai-conn',",
                "    optimization_mode => 'MINIMIZE_COST') AS is_defect,",
                "  AI.CLASSIFY(r.review_text,",
                "    categories => ['thermal_safety', 'electrical_failure',",
                "      'mechanical_defect', 'cosmetic_issue', 'not_a_defect'],",
                "    connection_id => 'us.vertex-ai-conn',",
                "    optimization_mode => 'MINIMIZE_COST') AS category,",
                "  AI.GENERATE_INT(('Rate urgency 1..5: ', r.review_text),",
                "    connection_id => 'us.vertex-ai-conn',",
                "    endpoint => 'gemini-2.5-flash').result AS urgency_1_to_5",
                "FROM `mfg_quality_demo.product_reviews` r;",
            ],
            "callouts": [
                {
                    "color": "blue",
                    "title": "🔵 AI.IF(prompt, ...) ➔ BOOL",
                    "text": "In: tupla (istruzione, review_text), connection_id, optimization_mode. Out: scalare BOOL diretto (is_defect).",
                },
                {
                    "color": "green",
                    "title": "🟢 AI.CLASSIFY(input, categories => [...]) ➔ STRING",
                    "text": "In: review_text, array categories (5 stringhe), connection_id, optimization_mode. Out: scalare STRING (category).",
                },
                {
                    "color": "yellow",
                    "title": "🟡 AI.GENERATE_INT(prompt, endpoint => ...) ➔ STRUCT",
                    "text": "In: tupla prompt, connection_id, endpoint. Out: STRUCT con campo .result tipizzato INT64 (urgency_1_to_5).",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 01 (LISTATO SQL):\n\n"
                "1. SINTASSI E PARAMETRI DI INGRESSO/USCITA DELLE 3 FUNZIONI AI:\n"
                "   • AI.IF(prompt, connection_id => ..., optimization_mode => ...):\n"
                "     - Input: tupla (istruzione, review_text), connection_id (opzionale), optimization_mode => 'MINIMIZE_COST' (opzionale).\n"
                "     - Output: scalare BOOL diretto (TRUE/FALSE, senza .result).\n"
                "   • AI.CLASSIFY(input, categories => [...], connection_id => ..., optimization_mode => ...):\n"
                "     - Input: review_text, array categories di 5 stringhe, connection_id (opzionale), optimization_mode (opzionale).\n"
                "     - Output: scalare STRING diretto contenente una delle categorie fornite.\n"
                "   • AI.GENERATE_INT(prompt, connection_id => ..., endpoint => ...):\n"
                "     - Input: tupla (istruzione, review_text), connection_id (opzionale), endpoint => 'gemini-2.5-flash'.\n"
                "     - Output: STRUCT con campo .result tipizzato INT64 (valutazione 1..5).\n\n"
                "2. È POSSIBILE OMETTERE connection_id => 'us.vertex-ai-conn' (CONNESSIONE DI DEFAULT / END-USER CREDENTIALS)?\n"
                "   • Sì: nelle funzioni BigQuery AI il parametro connection_id è OPZIONALE.\n"
                "   • Se si omette connection_id, BigQuery autentica la chiamata verso Vertex AI usando direttamente le End-User Credentials (EUC) dell'utente che esegue la query (purché l'utente abbia il ruolo IAM roles/aiplatform.user sul progetto).\n"
                "   • In alternativa, un amministratore può configurare una Cloud Resource Connection di default a livello di progetto o organizzazione (tramite ALTER PROJECT / ALTER ORGANIZATION SET OPTIONS) così da non doverla specificare in ogni query.\n"
                "   • Perché la passiamo esplicitamente negli script del workshop? Nei flussi di produzione (es. Stored Procedure schedulate come nel Modulo 03 o viste condivise) esplicitare connection_id delega i permessi Vertex AI al Service Account della connessione anziché dipendere dai permessi personali del singolo utente."
            ),
        },
    },
    {
        "id": "M2",
        "num": "02",
        "script": "sql/02_quick_triage.sql (Query 2b)",
        "name": "Classifica Semantica del Rischio Fisico",
        "subtitle": "AI.SCORE in ORDER BY per Prioritizzazione del Rischio Fisico",
        "goal": {
            "headline": "Ordinare le recensioni negative per pericolo fisico reale anziché solo per numero di stelle",
            "cards": [
                {
                    "color": "red",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "Tutte le review con 1 stella sembrano uguali",
                    "bullets": [
                        "Un graffio estetico e un cavo che prende fuoco possono avere entrambi 1 stella",
                        "Il semplice voto numerico non distingue un cliente insoddisfatto da un pericolo d'incendio",
                        "Il team Safety deve ispezionare per primi i casi di fumo, scintille o ustioni",
                    ],
                },
                {
                    "color": "blue",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Misurare il Pericolo Reale",
                    "bullets": [
                        "Quanto è grave il pericolo fisico descritto nel testo libero su una scala da 0 a 100?",
                        "Come evitare i pareggi (ex-aequo) tipici di una scala discreta 1..5 quando stabiliamo le priorità?",
                        "Quali sono le 15 segnalazioni più pericolose dell'intero catalogo da gestire subito?",
                    ],
                },
                {
                    "color": "green",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Coda di Priorità per il Team Safety",
                    "bullets": [
                        "Portare automaticamente in cima alla lista i reclami che descrivono rischi termici o elettrici gravi",
                        "Separare le vere emergenze di sicurezza dai reclami a 1–2 stelle legati a ritardi di consegna",
                        "Fornire una coda di ispezione immediata e ordinata per gravità",
                    ],
                },
            ],
            "kpi": "Punteggio Continuo di Rischio Fisico (0–100)  •  Zero Ex-Aequo in Classifica  •  Coda Top-15 Emergenze",
            "notes": "", 
        },
        "er": {
            "headline": "Punteggio di rischio con AI.SCORE",
            "postit": {
                "title": "📌 Filtro + AI.SCORE",
                "lines": [
                    "• Freccia SELECT verso la coda Top-15",
                    "• 1. Filtro relazionale: WHERE rating <= 2",
                    "  (dimezza le righe a costo zero)",
                    "• 2. Funzione: AI.SCORE(criterio, testo)",
                    "• Output: safety_hazard_score (FLOAT64)",
                    "• 3. ORDER BY score DESC LIMIT 15",
                ],
            },
            "note": "",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 02 (DIAGRAMMA E/R):\n"
                "• Anche qui la query esegue una SELECT con filtro e ordinamento: per questo usiamo la freccia semplice 'SELECT' (senza zampe di gallina E/R) "
                "dalla tabella product_reviews alla coda ordinata delle 15 recensioni più critiche."
            ),
        },
        "tech": {
            "headline": "Combinare il filtro relazionale SQL con lo scoring continuo gestito dall'LLM",
            "cards": [
                {
                    "color": "blue",
                    "tag": "PERCHÉ NON AI.GENERATE_INT?",
                    "title": "AI.SCORE vs AI.GENERATE_INT in ORDER BY",
                    "bullets": [
                        "Entrambe possono andare in ORDER BY, ma una scala intera 1..5 crea troppi ex-aequo",
                        "AI.SCORE genera automaticamente una rubrica di valutazione coerente sul criterio",
                        "Restituisce direttamente un FLOAT64 continuo (0–100) ideale per classifiche Top-K",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "OTTIMIZZAZIONE SQL + AI",
                    "title": "Filtro Preventivo (WHERE <= 2)",
                    "bullets": [
                        "Il motore SQL applica prima WHERE r.rating <= 2 a costo zero",
                        "AI.SCORE viene invocata solo sulle recensioni critiche (1 o 2 stelle)",
                        "Combinare filtri relazionali e funzioni AI dimezza latenza e token consumati",
                    ],
                },
                {
                    "color": "red",
                    "tag": "ESPRESSIONI SQL STANDARD",
                    "title": "Coda di Priorità Safety",
                    "bullets": [
                        "Lo scalare FLOAT64 di AI.SCORE viene ordinato con ORDER BY safety_hazard_score DESC",
                        "Permette di creare viste di allerta per i team di Sicurezza Prodotto",
                        "Rivela subito che i casi più gravi si concentrano sui lotti B-2025-01, 02 e 05",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 02 (PARTICOLARITÀ TECNICHE):\n"
                "• Approfondire il confronto AI.SCORE vs AI.GENERATE_INT: non è un limite sintattico di SQL (anche AI.GENERATE_INT(...).result può stare in ORDER BY), "
                "bensì una scelta di qualità del ranking (FLOAT64 continuo + rubrica auto-calibrata vs intero discreto senza rubrica).\n\n"
                "DOMANDA CHIAVE: SOLO AI.SCORE PUÒ ANDARE NELLA ORDER BY? PERCHÉ NON USARE AI.GENERATE_INT()?\n"
                "• Dal punto di vista sintattico SQL, QUALSIASI espressione scalare può andare nella ORDER BY, incluso AI.GENERATE_INT(...).result!\n"
                "• Tuttavia, per ordinare una coda di priorità (ranking Top-K), AI.SCORE è nettamente superiore a AI.GENERATE_INT per 3 motivi:\n"
                "  1. Granularità continua (FLOAT64 0–100) vs discreta (INT64 1–5): con AI.GENERATE_INT su scala 1..5, su 30 recensioni gravi ne avremmo "
                "decine tutte a pari merito (ex-aequo) con valore 5, rendendo arbitrario il taglio LIMIT 15. AI.SCORE restituisce un FLOAT64 continuo che separa finemente i casi.\n"
                "  2. Auto-calibrazione della rubrica (Auto-Rubric): AI.SCORE analizza il criterio di valutazione e genera internamente una rubrica di scoring "
                "strutturata e coerente applicata a tutte le righe della query.\n"
                "  3. Sintassi diretta managed: AI.SCORE restituisce direttamente un FLOAT64 (senza .result) e gestisce automaticamente il modello senza parametro endpoint."
            ),
        },
        "sql": {
            "headline": "Query 2b — Classifica delle 15 segnalazioni a maggior rischio fisico con AI.SCORE",
            "code": [
                "SELECT",
                "  r.review_id, r.sku, r.batch_id, r.rating,",
                "  AI.SCORE(",
                "    ('Urgency of physical safety hazard such as fire, smoke, ',",
                "     'melting plastic, electric shock, or burn risk: ',",
                "     r.review_text),",
                "    connection_id => 'us.vertex-ai-conn'",
                "  ) AS safety_hazard_score,",
                "  r.review_text",
                "FROM `mfg_quality_demo.product_reviews` r",
                "WHERE r.rating <= 2",
                "ORDER BY safety_hazard_score DESC",
                "LIMIT 15;",
            ],
            "callouts": [
                {
                    "color": "blue",
                    "title": "🔵 AI.SCORE(prompt, connection_id => ...) ➔ FLOAT64",
                    "text": "In: tupla (criterio di scoring, review_text) e connection_id. Out: scalare FLOAT64 diretto (0–100, senza .result).",
                },
                {
                    "color": "yellow",
                    "title": "🟡 WHERE r.rating <= 2",
                    "text": "In: predicato booleano sulla colonna INT64 rating applicato in ingresso sulla tabella product_reviews.",
                },
                {
                    "color": "red",
                    "title": "🔴 ORDER BY safety_hazard_score DESC LIMIT 15",
                    "text": "In: alias scalare FLOAT64 calcolato nella SELECT. Out: prime 15 righe ordinate per punteggio decrescente.",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 02 (LISTATO SQL):\n\n"
                "1. PARAMETRI E SINTASSI DI AI.SCORE:\n"
                "   • Input: tupla contenente (criterio_scoring, review_text) e parametro opzionale connection_id => 'us.vertex-ai-conn'.\n"
                "   • Output: scalare FLOAT64 continuo normalizzato tra 0 e 100 restituito direttamente (senza .result!).\n\n"
                "2. FILTRO E ORDINAMENTO IN SQL:\n"
                "   • WHERE r.rating <= 2 applica un filtro relazionale deterministico prima della SELECT per escludere recensioni neutre o positive.\n"
                "   • L'alias safety_hazard_score calcolato nella SELECT viene riutilizzato direttamente nella clausola ORDER BY DESC con LIMIT 15 per estrarre le anomalie fisiche prioritarie."
            ),
        },
    },
    {
        "id": "M3",
        "num": "03",
        "script": "sql/03_async_enrichment_pipeline.sql (Query 3a)",
        "name": "Pipeline di Arricchimento Incrementale",
        "subtitle": "AI.GENERATE con Output Strutturato + Anti-Join Incrementale + Auto-Retry Quota",
        "goal": {
            "headline": "Costruire un patrimonio informativo strutturato e continuo dalle recensioni in arrivo",
            "cards": [
                {
                    "color": "blue",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "Analisi Ripetitive e Costose",
                    "bullets": [
                        "Rianalizzare da zero l'intero storico recensioni a ogni report moltiplica tempi e costi",
                        "Ogni giorno arrivano nuove recensioni che devono essere integrate senza fermare i sistemi",
                        "Eventuali picchi di traffico o rallentamenti temporanei non devono far perdere dati",
                    ],
                },
                {
                    "color": "green",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Dal Testo Libero al Dato Strutturato",
                    "bullets": [
                        "Quali informazioni chiave (sentiment, difetto, categoria, componente, gravità, sintesi) servono per ogni recensione?",
                        "Come processare esclusivamente le recensioni nuove arrivate dall'ultimo aggiornamento?",
                        "Come garantire che eventuali errori temporanei vengano recuperati automaticamente?",
                    ],
                },
                {
                    "color": "red",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Scheda Tecnica per Ogni Recensione",
                    "bullets": [
                        "Trasformare ogni recensione testuale in una scheda strutturata pronta per i report aziendali",
                        "Individuare dal racconto del cliente quale specifico componente fisico ha ceduto",
                        "Disporre di un processo affidabile e aggiornabile continuamente sui nuovi arrivi",
                    ],
                },
            ],
            "kpi": "Estrazione Multi-Attributo Standardizzata  •  Elaborazione Solo dei Nuovi Arrivi  •  Continuità Operativa",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 03 (OBIETTIVO DI BUSINESS):\n"
                "• Qui passiamo dalla Fase 1 (esplorazione interattiva dei Moduli 01-02) alla Fase 2 (consolidamento di un patrimonio informativo strutturato permanente).\n"
                "• L'obiettivo di business è alimentare una pipeline incrementale resiliente che estragga automaticamente tutti i segnali di qualità dai reclami dei clienti senza duplicare il lavoro sui dati già processati."
            ),
        },
        "er": {
            "headline": "Left Anti-Join incrementale dalle recensioni grezze alla tabella review_insights",
            "postit": {
                "title": "📌 Pipeline & Resilienza",
                "lines": [
                    "• 1. LEFT JOIN su review_id WHERE IS NULL",
                    "  (seleziona solo recensioni nuove)",
                    "• 2. AI.GENERATE con output_schema",
                    "  (estrae 6 campi tipizzati in 1 chiamata)",
                    "• 3. Ciclo REPEAT ... UNTIL su ai_status:",
                    "  cancella e riprocessa gli errori HTTP 429",
                ],
            },
            "note": "",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 03 (DIAGRAMMA E/R):\n"
                "• Le relazioni nel diagramma mostrano le chiavi fisiche: 1:N sku tra products e product_reviews, e 1:1 review_id verso review_insights.\n"
                "• Logica chiave della Stored Procedure: Anti-Join incrementale (LEFT JOIN WHERE review_insights IS NULL), AI.GENERATE multi-campo con output_schema STRUCT a 6 campi e ciclo REPEAT per retry automatico su errori HTTP 429."
            ),
        },
        "tech": {
            "headline": "Tre pattern di ingegneria di produzione in un'unica Stored Procedure BigQuery",
            "cards": [
                {
                    "color": "green",
                    "tag": "1. OUTPUT STRUTTURATO",
                    "title": "output_schema Multi-Campo",
                    "bullets": [
                        "Un'unica chiamata AI.GENERATE restituisce una STRUCT SQL con 6 campi tipizzati",
                        "Più veloce ed economico rispetto a 6 funzioni AI separate sulla stessa riga",
                        "Vincola affected_component al vocabolario esatto della distinta base (BOM)",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "2. ANTI-JOIN INCREMENTALE",
                    "title": "LEFT JOIN ... WHERE IS NULL",
                    "bullets": [
                        "Confronta product_reviews con review_insights sulla chiave primaria review_id",
                        "Invia a Vertex AI esclusivamente le recensioni non ancora presenti a sistema",
                        "Rende la procedura enrich_new_reviews() idempotente e sicura da schedulare",
                    ],
                },
                {
                    "color": "red",
                    "tag": "3. RESILIENZA ALLE QUOTE",
                    "title": "Ciclo REPEAT + Filtro ai_status",
                    "bullets": [
                        "AI.GENERATE non fa fallire l'intera query se una riga riceve HTTP 429: scrive l'errore in .status",
                        "Il ciclo REPEAT elimina le righe con ai_status LIKE '%retryable error%'",
                        "L'anti-join successiva riprocessa automaticamente solo le righe fallite",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 03 (PARTICOLARITÀ TECNICHE):\n"
                "• Spiegare bene il comportamento di BigQuery AI sugli errori per-riga: se su 1000 righe 5 ricevono Resource Exhausted (429), "
                "la query NON fallisce ma popola il campo .status di quelle 5 righe con il messaggio di errore e mette i campi estratti a NULL.\n"
                "• Per questo salviamo ai.status nella colonna ai_status e usiamo REPEAT ... DELETE ... INSERT finché non restano 0 errori."
            ),
        },
        "sql": {
            "headline": "Query 3a — Stored Procedure incrementale con output_schema e recupero errori 429",
            "code": [
                "CREATE OR REPLACE PROCEDURE `mfg_quality_demo.enrich_new_reviews`()",
                "BEGIN",
                "  REPEAT",
                "    DELETE FROM `mfg_quality_demo.review_insights`",
                "    WHERE ai_status LIKE '%A retryable error occurred%';",
                "    INSERT INTO `mfg_quality_demo.review_insights`",
                "    SELECT review_id, ai.sentiment, ai.is_defect_report,",
                "      ai.defect_category, ai.affected_component,",
                "      ai.severity, ai.summary, ai.status, CURRENT_TIMESTAMP()",
                "    FROM (",
                "      SELECT r.review_id, AI.GENERATE(",
                "        ('Analyze ', p.product_name, ': ', r.review_text),",
                "        connection_id => 'us.vertex-ai-conn',",
                "        endpoint => 'gemini-2.5-flash',",
                "        output_schema => 'sentiment STRING, is_defect_report BOOL,",
                "          defect_category STRING, affected_component STRING,",
                "          severity INT64, summary STRING') AS ai",
                "      FROM `mfg_quality_demo.product_reviews` r",
                "      JOIN `mfg_quality_demo.products` p USING (sku)",
                "      LEFT JOIN `mfg_quality_demo.review_insights` i USING (review_id)",
                "      WHERE i.review_id IS NULL);",
                "  UNTIL (SELECT COUNT(*) FROM `mfg_quality_demo.review_insights`",
                "         WHERE ai_status LIKE '%retryable error%') = 0 END REPEAT;",
                "END;",
            ],
            "callouts": [
                {
                    "color": "green",
                    "title": "🟢 AI.GENERATE(..., output_schema => '...') ➔ STRUCT",
                    "text": "In: tupla prompt, connection_id, endpoint, output_schema DDL. Out: STRUCT ai con 6 campi tipizzati + ai.status (STRING).",
                },
                {
                    "color": "yellow",
                    "title": "🟡 LEFT JOIN ... USING (review_id) WHERE i.review_id IS NULL",
                    "text": "In: chiave review_id tra product_reviews (r) e review_insights (i). Out: sole righe di r con i.review_id NULL.",
                },
                {
                    "color": "red",
                    "title": "🔴 REPEAT ... DELETE ... INSERT ... UNTIL count = 0",
                    "text": "In: blocco procedurale SQL con DELETE su ai_status LIKE '%retryable error%' e condizione scalare booleana UNTIL.",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 03 (LISTATO SQL):\n\n"
                "1. PARAMETRI E STRUTTURA DI OUTPUT DI AI.GENERATE:\n"
                "   • Input: tupla prompt (istruzioni di estrazione multi-campo), connection_id => 'us.vertex-ai-conn', endpoint => 'gemini-2.5-flash' e parametro DDL output_schema.\n"
                "   • Output: STRUCT ai con 6 campi tipizzati (is_defect_report BOOL, category STRING, severity INT64, affected_component STRING, failure_mode STRING, root_cause_hypothesis STRING) più il campo di stato ai.status (STRING).\n\n"
                "2. SINTASSI ANTI-JOIN E PROCEDURA DI RETRY:\n"
                "   • LEFT JOIN product_reviews r LEFT JOIN review_insights i USING (review_id) WHERE i.review_id IS NULL: pattern standard per filtrare in ingresso solo le righe nuove.\n"
                "   • Blocco procedurale REPEAT ... UNTIL: cancella le righe andate in errore temporaneo (DELETE FROM review_insights WHERE ai_status LIKE '%retryable error%') e ripete la INSERT incrementale finché non restano fallimenti."
            ),
        },
    },
    {
        "id": "M4",
        "num": "04",
        "script": "sql/04_root_cause_analysis.sql (Query 4a + 4b + 4c)",
        "name": "Analisi della Causa Radice sulla Supply Chain",
        "subtitle": "Join Relazionale Ibrida AI-Distinta Base: da Testo Libero a Lotto HE-4471 e Macchina CAL-02",
        "goal": {
            "headline": "Risalire dai sintomi descritti dai clienti al lotto di componenti e al macchinario responsabile",
            "cards": [
                {
                    "color": "blue",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "Il Cliente Non Conosce la Fabbrica",
                    "bullets": [
                        "Il cliente descrive il sintomo ('non si spegne mai'), ma non sa quale componente interno sia guasto",
                        "Ogni lotto di prodotto finito assembla 5 componenti diversi provenienti da più fornitori",
                        "Un richiamo generalizzato dell'intero catalogo costerebbe milioni e danneggerebbe il brand",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Fornitore o Macchinario di Linea?",
                    "bullets": [
                        "Quale specifico lotto di componenti registra la percentuale più alta di reclami sul campo?",
                        "Il difetto dipende dalla materia prima del fornitore o dal macchinario che lo ha installato?",
                        "Altri lotti dello stesso fornitore montati su linee diverse presentano lo stesso problema?",
                    ],
                },
                {
                    "color": "red",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Richiamo Chirurgico e Verifica",
                    "bullets": [
                        "Isolare con precisione il singolo lotto difettoso e i soli lotti di produzione impattati",
                        "Accertare con dati oggettivi se la responsabilità è del fornitore o di una macchina di stabilimento",
                        "Intervenire solo sulla linea e sui lotti coinvolti risparmiando milioni in richiami inutili",
                    ],
                },
            ],
            "kpi": "Tracciabilità Sintomo ➔ Componente ➔ Lotto ➔ Macchina  •  Verifica Responsabilità Fornitore vs Fabbrica",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 04 (OBIETTIVO DI BUSINESS):\n"
                "• Questo è il cuore investigativo del workshop: partire dal sintomo testuale del cliente per capire se il guasto nasce dal fornitore "
                "del componente o da un macchinario di assemblaggio in fabbrica."
            ),
        },
        "er": {
            "headline": "Attraversamento di 5 tabelle della supply chain guidato da affected_component = component_type",
            "postit": {
                "title": "📌 Ponte Semantico-BOM",
                "lines": [
                    "• Tabella ponte M:N batch_components:",
                    "  PK composta (batch_id, lot_id)",
                    "• Ogni batch monta 5 componenti diversi",
                    "• Condizione chiave di disambiguazione:",
                    "  i.affected_component = cl.component_type",
                    "• Isola solo il componente citato dal cliente!",
                ],
            },
            "note": "",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 04 (DIAGRAMMA E/R):\n"
                "• Domanda frequente degli studenti: 'Qual è la Primary Key di batch_components?' Risposta: è una tabella di appoggio associativa M:N "
                "la cui chiave primaria è la chiave composta (batch_id, lot_id), dove ciascun attributo è anche Foreign Key.\n"
                "• Senza il predicato semantico i.affected_component = cl.component_type, una recensione sul termostato nel batch B-2025-01 "
                "verrebbe associata anche al coperchio, alla guarnizione e al cavo di quello stesso batch!"
            ),
        },
        "tech": {
            "headline": "Perché il predicato ibrido semantico-relazionale è il fulcro dell'intera architettura",
            "cards": [
                {
                    "color": "green",
                    "tag": "IL PREDICATO CHIAVE",
                    "title": "Disambiguare la Distinta Base M:N",
                    "bullets": [
                        "Un batch (es. B-2025-01) contiene 5 componenti: resistenza, sensore, coperchio, guarnizione, cavo",
                        "La JOIN tradizionale su batch_id assocerebbe il reclamo termico anche al coperchio!",
                        "Il filtro i.affected_component = cl.component_type isola solo il componente citato dal cliente",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "TASSO DI IMPLICAZIONE",
                    "title": "COUNT(DISTINCT CASE WHEN ...)",
                    "bullets": [
                        "Denominatore: tutte le recensioni dei batch che montano quel lot_id",
                        "Numeratore: solo le recensioni difettose in cui l'AI ha indicato quel component_type",
                        "HE-4471 raggiunge 17 / 22 = 77.3% con severità media 4.5 / 5.0",
                    ],
                },
                {
                    "color": "red",
                    "tag": "VERIFICA INCROCIATA",
                    "title": "Scagionare il Fornitore ThermoCore",
                    "bullets": [
                        "Query 4b aggrega su machines: CAL-02 emerge con manutenzione ferma al 2025-01-17",
                        "Query 4c confronta tutti i heating_element: HE-4472 (stesso fornitore ThermoCore ma su CAL-01) è al 9%",
                        "Dimostra matematicamente che il problema è la taratura di CAL-02, non la materia prima",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 04 (PARTICOLARITÀ TECNICHE):\n"
                "• Far notare come l'output dell'LLM (affected_component) diventi a tutti gli effetti una chiave di join relazionale."
            ),
        },
        "sql": {
            "headline": "Query 4a & 4b — Vista difettosità per lotto e risalita alla calibratrice CAL-02",
            "code": [
                "CREATE OR REPLACE VIEW `mfg_quality_demo.v_component_lot_defect_rates` AS",
                "SELECT cl.lot_id, cl.component_type, cl.supplier, cl.installed_by_machine_id,",
                "  COUNT(DISTINCT r.review_id) AS total_reviews_for_lot,",
                "  COUNT(DISTINCT CASE WHEN i.is_defect_report",
                "    AND i.affected_component = cl.component_type",
                "    THEN r.review_id END) AS defect_reviews,",
                "  ROUND(SAFE_DIVIDE(COUNT(DISTINCT CASE WHEN i.is_defect_report",
                "    AND i.affected_component = cl.component_type THEN r.review_id END),",
                "    COUNT(DISTINCT r.review_id)), 2) AS implication_rate",
                "FROM `mfg_quality_demo.component_lots` cl",
                "JOIN `mfg_quality_demo.batch_components` bc USING (lot_id)",
                "JOIN `mfg_quality_demo.product_reviews` r   USING (batch_id)",
                "JOIN `mfg_quality_demo.review_insights` i   USING (review_id)",
                "GROUP BY 1, 2, 3, 4;",
                "-- 4b. Roll-up per Macchina di Fabbrica:",
                "SELECT m.machine_id, m.last_maintenance_date,",
                "  SUM(v.defect_reviews) AS defects, ROUND(AVG(v.implication_rate), 2) AS rate",
                "FROM `mfg_quality_demo.v_component_lot_defect_rates` v",
                "JOIN `mfg_quality_demo.machines` m ON m.machine_id = v.installed_by_machine_id",
                "GROUP BY 1, 2 ORDER BY defects DESC;",
            ],
            "callouts": [
                {
                    "color": "green",
                    "title": "🟢 COUNT(DISTINCT CASE WHEN ... THEN r.review_id END)",
                    "text": "In: predicato i.is_defect_report AND i.affected_component = cl.component_type. Out: conteggio INT64 (defect_reviews).",
                },
                {
                    "color": "yellow",
                    "title": "🟡 SAFE_DIVIDE(numeratore, denominatore) ➔ FLOAT64",
                    "text": "In: defect_reviews (INT64) e total_reviews_for_lot (INT64). Out: rapporto FLOAT64 arrotondato con ROUND(..., 2).",
                },
                {
                    "color": "red",
                    "title": "🔴 JOIN ... USING (lot_id, batch_id, review_id)",
                    "text": "In: chiavi di join tra le 4 tabelle e GROUP BY 1, 2, 3, 4. Out: vista per lot_id e query di roll-up su machine_id.",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 04 (LISTATO SQL):\n\n"
                "1. PARAMETRI E AGGREGAZIONI CONDIZIONALI:\n"
                "   • COUNT(DISTINCT CASE WHEN i.is_defect_report AND i.affected_component = cl.component_type THEN r.review_id END): conteggio INT64 dei soli reclami pertinenti al componente del lotto.\n"
                "   • SAFE_DIVIDE(defect_reviews, total_reviews_for_lot): rapporto matematico FLOAT64 protetto da divisioni per zero, formattato con ROUND(..., 2) come defect_rate.\n\n"
                "2. CHIAVI DI JOIN E LIVELLI DI AGGREGAZIONE:\n"
                "   • La prima query raggruppa con GROUP BY 1, 2, 3, 4 per calcolare la vista v_component_lot_defect_rates a livello di singolo lot_id.\n"
                "   • La seconda query esegue una join con machines m USING (machine_id) e calcola il roll-up aggregato per macchinario d'assemblaggio."
            ),
        },
    },
    {
        "id": "M5",
        "num": "05",
        "script": "sql/05_executive_summary.sql (Query 5)",
        "name": "Piano d'Azione Esecutivo da SQL a Prompt",
        "subtitle": "Pre-Aggregazione Deterministica SQL + Sintesi AI.GENERATE in Singola Chiamata",
        "goal": {
            "headline": "Sintetizzare le evidenze numeriche di fabbrica in un piano d'azione esecutivo per la direzione",
            "cards": [
                {
                    "color": "blue",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "Dai Dati Tecnici alla Decisione",
                    "bullets": [
                        "La direzione operativa (VP of Manufacturing) non ha tempo di interpretare tabelle e metriche grezze",
                        "Un riassunto generico scritto senza basi numeriche certificate rischia di contenere imprecisioni",
                        "In una crisi di qualità serve un documento decisionale immediato, sintetico e verificabile",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Cosa Deve Decidere il Board Oggi?",
                    "bullets": [
                        "Qual è la causa radice accertata e quali evidenze numeriche giustificano tale conclusione?",
                        "Quali azioni immediate vanno intraprese in fabbrica, in magazzino e verso i fornitori?",
                        "Qual è l'impatto stimato sui clienti finali e quali lotti vanno messi in quarantena?",
                    ],
                },
                {
                    "color": "green",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Briefing Esecutivo Certificato",
                    "bullets": [
                        "Produrre una nota direzionale di 250 parole ancorata al 100% alle metriche reali di fabbrica",
                        "Articolare chiaramente Causa Radice, Azioni Immediate e Impatto Clienti",
                        "Azzerare il rischio di cifre inventate o contraddittorie nel passaggio tra analisti e management",
                    ],
                },
            ],
            "kpi": "Comunicazione Direzionale in 3 Punti  •  100% Ancorata alle Metriche Certificate  •  Pronta per il Board",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 05 (OBIETTIVO DI BUSINESS):\n"
                "• Regola d'oro: non chiedere mai a un LLM di fare conteggi o divisioni su decine di righe grezze. "
                "Fai calcolare i numeri esatti a BigQuery SQL e passa a Gemini solo la tabella riassuntiva."
            ),
        },
        "er": {
            "headline": "Compressione delle viste multi-riga in una CTE a riga singola tramite STRING_AGG",
            "postit": {
                "title": "📌 Pattern SQL-to-Prompt",
                "lines": [
                    "• 1. FORMAT(...) formatta i Top 5 lotti",
                    "• 2. STRING_AGG(..., '\\n') collassa le",
                    "  5 righe in 1 sola riga (top_lots)",
                    "• 3. AI.GENERATE viene invocata",
                    "  esattamente 1 volta (1 riga = 1 chiamata)",
                    "• 4. Output: executive_briefing (STRING)",
                ],
            },
            "note": "",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 05 (DIAGRAMMA E/R):\n"
                "• Il diagramma mostra il join 1:N su machine_id tra machines e v_component_lot_defect_rates.\n"
                "• Pattern anti-allucinazione: STRING_AGG comprime le N righe calcolate dal DB in 1 singolo scalare prima di invocare AI.GENERATE, separando il calcolo deterministico dalla sintesi verbale."
            ),
        },
        "tech": {
            "headline": "Il pattern anti-allucinazione 'SQL-to-Prompt' per il reporting direzionale",
            "cards": [
                {
                    "color": "yellow",
                    "tag": "SEPARAZIONE DEI COMPITI",
                    "title": "SQL per la Matematica, LLM per la Sintesi",
                    "bullets": [
                        "Gli LLM non devono mai calcolare aggregazioni o percentuali su decine di righe",
                        "Il motore SQL di BigQuery calcola in modo deterministico tassi, conteggi e date",
                        "L'LLM riceve solo l'evidenza già aggregata e certificata",
                    ],
                },
                {
                    "color": "green",
                    "tag": "COMPRESSIONE DEL CONTESTO",
                    "title": "FORMAT + STRING_AGG(..., '\\n')",
                    "bullets": [
                        "FORMAT costruisce una riga testuale strutturata per ciascuno dei 5 lotti principali",
                        "STRING_AGG collassa le 5 righe in un unico scalare di testo (finding_lines)",
                        "Riduce l'input della query finale a 1 sola riga = 1 sola invocazione Gemini",
                    ],
                },
                {
                    "color": "blue",
                    "tag": "PROMPT STRUTTURATO",
                    "title": "Output Azionabile in 3 Sezioni",
                    "bullets": [
                        "Il prompt richiede esplicitamente: (1) Causa Radice, (2) Azioni Immediate, (3) Impatto Clienti",
                        "Cita i numeri esatti (77% vs 9%) per giustificare lo scagionamento del fornitore",
                        "Pronto per essere inviato via email o mostrato in una dashboard Looker",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 05 (PARTICOLARITÀ TECNICHE):\n"
                "• Confrontare questo approccio (STRING_AGG manuale + AI.GENERATE) con AI.AGG che vedremo nel Modulo 07."
            ),
        },
        "sql": {
            "headline": "Query 5 — Pre-aggregazione deterministica in CTE e generazione del Piano d'Azione",
            "code": [
                "WITH top_lots AS (",
                "  SELECT STRING_AGG(FORMAT(",
                "    'Lot %s (%s, %s, machine=%s, maint=%t): %d/%d (%.0f%%)',",
                "    v.lot_id, v.component_type, v.supplier,",
                "    v.installed_by_machine_id, m.last_maintenance_date,",
                "    v.defect_reviews, v.total_reviews_for_lot,",
                "    v.implication_rate * 100",
                "  ), '\\n' ORDER BY v.defect_reviews DESC) AS finding_lines",
                "  FROM (",
                "    SELECT * FROM `mfg_quality_demo.v_component_lot_defect_rates`",
                "    ORDER BY defect_reviews DESC LIMIT 5",
                "  ) v JOIN `mfg_quality_demo.machines` m",
                "    ON m.machine_id = v.installed_by_machine_id",
                ")",
                "SELECT AI.GENERATE(",
                "  ('Write a 250-word executive briefing with: (1) Root Cause, ',",
                "   '(2) Immediate Actions, (3) Customer Impact.\\n', t.finding_lines),",
                "  connection_id => 'us.vertex-ai-conn',",
                "  endpoint => 'gemini-2.5-flash').result AS executive_briefing",
                "FROM top_lots t;",
            ],
            "callouts": [
                {
                    "color": "yellow",
                    "title": "🟡 FORMAT('Lot %s ...', ...) ➔ STRING",
                    "text": "In: stringa di formato (%s, %t, %d, %.0f%%) e 8 colonne tipizzate. Out: 1 valore STRING formattato per riga.",
                },
                {
                    "color": "red",
                    "title": "🔴 STRING_AGG(..., '\\n' ORDER BY ...) ➔ STRING",
                    "text": "In: espressione STRING per riga e separatore '\\n'. Out: 1 sola riga con colonna finding_lines (STRING) nella CTE top_lots.",
                },
                {
                    "color": "blue",
                    "title": "🔵 AI.GENERATE((prompt, t.finding_lines), ...).result",
                    "text": "In: tupla (istruzione, finding_lines), connection_id, endpoint. Out: STRUCT con campo .result STRING (executive_briefing).",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 05 (LISTATO SQL):\n\n"
                "1. SERIALIZZAZIONE E COMPRESSIONE IN SQL:\n"
                "   • FORMAT('Lot %s (%s, %s): installed by %s on %s...', ...): serializza 8 colonne eterogenee (%s stringa, %t data, %d intero, %.0f%% percentuale) in un'unica riga testuale coerente.\n"
                "   • STRING_AGG(..., '\\n' ORDER BY defect_rate DESC, total_defects DESC): funzione di aggregazione che comprime N righe in un unico scalare STRING (finding_lines).\n\n"
                "2. PARAMETRI DI AI.GENERATE:\n"
                "   • Input: tupla contenente l'istruzione direzionale del prompt e il blocco testuale t.finding_lines, connection_id => 'us.vertex-ai-conn', endpoint => 'gemini-2.5-flash'.\n"
                "   • Output: STRUCT con campo .result di tipo STRING contenente l'executive briefing generato."
            ),
        },
    },
    {
        "id": "M6",
        "num": "06",
        "script": "sql/07_key_drivers_and_ai_agg_enhancements.sql (Query 7a)",
        "name": "Scoperta Automatica delle Cause Radice",
        "subtitle": "AI.KEY_DRIVERS per l'Analisi Automatica Multi-Dimensionale dei Fattori Critici",
        "goal": {
            "headline": "Individuare automaticamente quale combinazione di fattori spiega l'anomalia di qualità",
            "cards": [
                {
                    "color": "blue",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "E se Avessimo 50 Dimensioni?",
                    "bullets": [
                        "In uno stabilimento reale esistono decine di variabili (turno, linea, fornitore, lotto, macchina, SKU...)",
                        "Verificare manualmente ogni singola ipotesi o incrocio richiede giorni di analisi per tentativi",
                        "Le cause radice più insidiose nascono spesso dall'interazione tra più fattori contemporaneamente",
                    ],
                },
                {
                    "color": "green",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Quali Fattori Spingono i Difetti?",
                    "bullets": [
                        "Quale attributo singolo o combinazione di attributi concentra un numero anomalo di difetti?",
                        "Quanto è più frequente il guasto in quel segmento rispetto alla media normale di fabbrica?",
                        "Quali combinazioni hanno il maggiore impatto complessivo sul totale dei reclami?",
                    ],
                },
                {
                    "color": "red",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Diagnosi Multi-Fattore Istantanea",
                    "bullets": [
                        "Esplorare in automatico tutte le combinazioni possibili della catena produttiva senza ipotesi a priori",
                        "Ordinare i fattori sospetti per rilevanza statistica e volume di difetti spiegati",
                        "Ridurre da giorni a pochi secondi il tempo necessario per isolare un'anomalia complessa",
                    ],
                },
            ],
            "kpi": "Esplorazione Automatica Multi-Dimensionale  •  Zero Ipotesi Manuali a Priori  •  Ranking per Impatto Reale",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 06 (OBIETTIVO DI BUSINESS):\n"
                "• Mostrare il problema di scalabilità dell'analista umano quando i fattori della supply chain sono decine.\n"
                "• L'obiettivo di business è isolare automaticamente e in modo statisticamente rigoroso quali combinazioni di fornitori, macchinari e lotti generano un aumento anomalo del tasso di difettosità."
            ),
        },
        "er": {
            "headline": "Join denormalizzata a 6 tabelle passata in input alla funzione tabellare AI.KEY_DRIVERS",
            "postit": {
                "title": "📌 AI.KEY_DRIVERS",
                "lines": [
                    "• Funzione Tabellare (TVF) in FROM",
                    "• Input: Star Join a 6 tabelle con",
                    "  6 dimensioni + target booleano",
                    "• Target: is_implicated_defect (BOOL)",
                    "• Output: combinazioni key_driver",
                    "  ordinate per impact_score e lift",
                ],
            },
            "note": "",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 06 (DIAGRAMMA E/R):\n"
                "• Tutte le relazioni nel diagramma usano le chiavi fisiche (review_id, sku, batch_id, lot_id) con PK composta su batch_components.\n"
                "• Funzionamento: l'intera join denormalizzata a 5 tabelle viene passata dentro TABLE(...) a AI.KEY_DRIVERS, che calcola automaticamente combinazioni di fattori e Lift statistico rispetto al target."
            ),
        },
        "tech": {
            "headline": "Come interpretare le metriche statistiche di contributo restituite da AI.KEY_DRIVERS",
            "cards": [
                {
                    "color": "blue",
                    "tag": "COME FUNZIONA AI.KEY_DRIVERS",
                    "title": "Funzione Tabellare (TVF)",
                    "bullets": [
                        "Prende in input una TABLE(SELECT ...) con le dimensioni candidate e un target booleano",
                        "Valuta sia i singoli attributi (es. sku='KTL-100') sia le intersezioni multi-attributo",
                        "Filtra il rumore statistico e restituisce solo i fattori significativi",
                    ],
                },
                {
                    "color": "green",
                    "tag": "LE METRICHE STATISTICHE",
                    "title": "Lift, Odds Ratio e Impatto",
                    "bullets": [
                        "lift: quanto è più probabile il difetto in quel segmento rispetto alla media globale (16x)",
                        "unexpected_True_ratio: quota di difetti 'in eccesso' spiegata da quel segmento (+69%)",
                        "impact_score: combina volume di difetti e lift per dare una priorità unica",
                    ],
                },
                {
                    "color": "red",
                    "tag": "CONFRONTO DIDATTICO",
                    "title": "Modulo 04 vs Modulo 06",
                    "bullets": [
                        "Modulo 04 (SQL classico): ottimo quando conosci già la gerarchia da verificare",
                        "Modulo 06 (AI.KEY_DRIVERS): imbattibile per scoprire cause radice al buio su decine di colonne",
                        "Entrambi convergono esattamente su HE-4471 + ThermoCore + CAL-02",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 06 (PARTICOLARITÀ TECNICHE):\n"
                "• Spiegare che AI.KEY_DRIVERS non usa un LLM generativo (non richiede connection_id Vertex AI!), ma il motore statistico di Contribution Analysis integrato in BigQuery."
            ),
        },
        "sql": {
            "headline": "Query 7a — Scoperta automatica dei fattori di difettosità con AI.KEY_DRIVERS",
            "code": [
                "SELECT key_driver, true_count AS defect_matches, total_count,",
                "  ROUND(lift, 2) AS lift, ROUND(odds_ratio, 2) AS odds_ratio,",
                "  ROUND(unexpected_True_ratio, 2) AS unexpected_defect_ratio,",
                "  ROUND(impact_score, 2) AS impact_score",
                "FROM AI.KEY_DRIVERS(",
                "  TABLE(",
                "    SELECT p.sku, pb.production_line, cl.component_type,",
                "      cl.supplier, cl.installed_by_machine_id, cl.lot_id,",
                "      (i.is_defect_report AND",
                "       i.affected_component = cl.component_type) AS is_implicated_defect",
                "    FROM `mfg_quality_demo.review_insights` i",
                "    JOIN `mfg_quality_demo.product_reviews` r     USING (review_id)",
                "    JOIN `mfg_quality_demo.products` p            USING (sku)",
                "    JOIN `mfg_quality_demo.production_batches` pb USING (batch_id)",
                "    JOIN `mfg_quality_demo.batch_components` bc   USING (batch_id)",
                "    JOIN `mfg_quality_demo.component_lots` cl     USING (lot_id)",
                "  ), is_implicated_defect",
                ") ORDER BY impact_score DESC LIMIT 10;",
            ],
            "callouts": [
                {
                    "color": "blue",
                    "title": "🔵 FROM AI.KEY_DRIVERS(TABLE(SELECT ...), target)",
                    "text": "In (Arg 1): TABLE(...) contenente 6 colonne dimensione (STRING) e 1 colonna target (BOOL). Sintassi TVF in FROM.",
                },
                {
                    "color": "green",
                    "title": "🟢 Arg 2: is_implicated_defect (BOOL)",
                    "text": "In (Arg 2): identificatore della colonna booleana nella TABLE di input rispetto a cui calcolare le metriche.",
                },
                {
                    "color": "red",
                    "title": "🔴 Colonne di Output della TVF",
                    "text": "Out: key_driver (ARRAY<STRING>), true_count (INT64), total_count (INT64), lift, odds_ratio, unexpected_True_ratio, impact_score (FLOAT64).",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 06 (LISTATO SQL):\n\n"
                "1. SINTASSI DELLA TABLE-VALUED FUNCTION (TVF):\n"
                "   • FROM AI.KEY_DRIVERS(TABLE(subquery), is_implicated_defect): accetta una subquery denormalizzata incapsulata nel costruttore TABLE(...) e il nome della colonna target booleana.\n"
                "   • La subquery deve includere sia le dimensioni categoriche (STRING) sia la colonna target booleana (BOOL).\n\n"
                "2. SCHEMA DI OUTPUT DELLA TVF:\n"
                "   • key_driver: ARRAY<STRING> contenente le combinazioni di attributi (es. ['component_type=heating_element', 'installed_by_machine_id=CAL-02']).\n"
                "   • true_count, total_count: INT64 con frequenze osservate.\n"
                "   • lift, odds_ratio, unexpected_True_ratio, impact_score: FLOAT64 con metriche di significatività statistica e impatto sul totale."
            ),
        },
    },
    {
        "id": "M7",
        "num": "07",
        "script": "sql/07_key_drivers_and_ai_agg_enhancements.sql (Query 7b)",
        "name": "Aggregazione Semantica Multi-Riga Nativa",
        "subtitle": "AI.AGG dentro GROUP BY per la Sintesi Automatica per Prodotto (SKU)",
        "goal": {
            "headline": "Fornire a ogni Product Manager un quadro sintetico di KPI e sintomi ricorrenti per prodotto",
            "cards": [
                {
                    "color": "blue",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "Un Report per Ogni Product Manager",
                    "bullets": [
                        "Ogni Product Manager è responsabile di una specifica linea di prodotto (Bollitore, Frullatore, Friggitrice...)",
                        "I numeri da soli (# reclami, gravità media) dicono quanto è grande il problema, ma non cosa succede",
                        "Leggere centinaia di recensioni per capire i sintomi ricorrenti di ogni prodotto richiede troppo tempo",
                    ],
                },
                {
                    "color": "green",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Numeri e Sintesi nella Stessa Vista",
                    "bullets": [
                        "Quante segnalazioni di difetto e quale gravità media registra ciascun prodotto a catalogo?",
                        "Quali sono i sintomi ricorrenti e i lotti più citati dai clienti per quello specifico prodotto?",
                        "Quale verifica ingegneristica prioritaria emerge dai reclami di ciascuna linea?",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Cruscotto Qualità per Prodotto",
                    "bullets": [
                        "Affiancare in un'unica tabella gli indicatori quantitativi e il riassunto qualitativo dei guasti per prodotto",
                        "Distinguere subito i prodotti con veri problemi strutturali da quelli con sole segnalazioni estetiche",
                        "Offrire a ciascun responsabile di prodotto indicazioni operative immediate",
                    ],
                },
            ],
            "kpi": "Vista Unificata per Prodotto (SKU)  •  KPI Quantitativi + Sintesi Qualitativa dei Sintomi  •  Priorità per i PM",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 07 (OBIETTIVO DI BUSINESS):\n"
                "• L'obiettivo di business è sintetizzare grandi volumi di reclami testuali a livello di singolo prodotto (SKU).\n"
                "• Vogliamo produrre un riepilogo qualitativo azionabile per il team di ingegneria di prodotto senza dover leggere manualmente centinaia di recensioni per ciascun modello."
            ),
        },
        "er": {
            "headline": "Aggregazione semantica molti-a-uno sulle recensioni raggruppate per products.sku",
            "postit": {
                "title": "📌 Funzione AI.AGG",
                "lines": [
                    "• Raggruppamento: GROUP BY p.sku",
                    "• Aggregazioni SQL: COUNT(*), AVG()",
                    "• Aggregazione AI: AI.AGG(campi_riga,",
                    "  instructions => 'Summarize...')",
                    "• Riduce N recensioni per SKU in",
                    "  1 sintesi testuale per prodotto",
                ],
            },
            "note": "",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 07 (DIAGRAMMA E/R):\n"
                "• Il diagramma E/R mostra le 3 tabelle fisiche coinvolte con le rispettive chiavi (1:1 review_id e 1:N sku).\n"
                "• Funzionamento di AI.AGG: opera direttamente sul raggruppamento GROUP BY p.sku riducendo N recensioni in 1 sintesi testuale per ciascun prodotto."
            ),
        },
        "tech": {
            "headline": "AI.GENERATE scalare vs AI.AGG aggregata: quando usare l'una o l'altra",
            "cards": [
                {
                    "color": "blue",
                    "tag": "FUNZIONE AGGREGATA NATIVA",
                    "title": "AI.AGG vs AI.GENERATE",
                    "bullets": [
                        "AI.GENERATE è una funzione scalare (1 riga in input ➔ 1 riga in output)",
                        "AI.AGG è una funzione di aggregazione (N righe di un gruppo ➔ 1 valore per gruppo)",
                        "Si scrive direttamente nella SELECT insieme a COUNT(*), AVG(), MAX() con GROUP BY",
                    ],
                },
                {
                    "color": "green",
                    "tag": "SCALABILITÀ GERARCHICA",
                    "title": "Gestione Automatica del Contesto",
                    "bullets": [
                        "Con STRING_AGG manuale, se un gruppo ha 50.000 recensioni la stringa esplode",
                        "AI.AGG gestisce nativamente l'aggregazione di grandi insiemi di righe per gruppo",
                        "Accetta una tupla con i campi di ogni riga (es. batch_id, affected_component, review_text)",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "SEPARAZIONE ISTRUZIONE / DATI",
                    "title": "Parametro instructions => ...",
                    "bullets": [
                        "Il primo argomento di AI.AGG contiene solo l'espressione riga-per-riga da aggregare",
                        "Il parametro nominale instructions => '...' definisce il task di sintesi a livello di gruppo",
                        "Evita di ripetere il testo del prompt per ogni singola riga del gruppo!",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 07 (PARTICOLARITÀ TECNICHE):\n"
                "• Sottolineare il vantaggio del parametro instructions => '...' rispetto a concatenare il prompt in ogni riga."
            ),
        },
        "sql": {
            "headline": "Query 7b — Sintesi multi-riga per prodotto usando AI.AGG dentro GROUP BY",
            "code": [
                "SELECT",
                "  p.sku, p.product_name,",
                "  COUNT(*) AS defect_review_count,",
                "  ROUND(AVG(i.severity), 1) AS avg_severity,",
                "  AI.AGG(",
                "    ('Batch ', r.batch_id, ' [', i.affected_component, ']: ',",
                "     r.review_text),",
                "    instructions => 'Summarize recurring defect symptoms, '",
                "      || 'affected batches, and engineering check in 2 bullets.',",
                "    connection_id => 'us.vertex-ai-conn',",
                "    endpoint => 'gemini-2.5-flash'",
                "  ) AS product_defect_synthesis",
                "FROM `mfg_quality_demo.review_insights` i",
                "JOIN `mfg_quality_demo.product_reviews` r USING (review_id)",
                "JOIN `mfg_quality_demo.products` p        USING (sku)",
                "WHERE i.is_defect_report = TRUE",
                "GROUP BY p.sku, p.product_name",
                "ORDER BY defect_review_count DESC;",
            ],
            "callouts": [
                {
                    "color": "blue",
                    "title": "🔵 AI.AGG(valore_riga, ...) ➔ STRING",
                    "text": "In (Arg 1): tupla di colonne per riga (batch_id, affected_component, review_text). Out: scalare STRING diretto (senza .result).",
                },
                {
                    "color": "green",
                    "title": "🟢 instructions => '...'",
                    "text": "In (Parametri nominali): instructions (STRING), connection_id (STRING), endpoint (STRING).",
                },
                {
                    "color": "yellow",
                    "title": "🟡 GROUP BY p.sku, p.product_name",
                    "text": "In: chiavi di raggruppamento SQL. Out: 1 riga per SKU con COUNT(*) (INT64), AVG(severity) (FLOAT64) e AI.AGG (STRING).",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 07 (LISTATO SQL):\n\n"
                "1. PARAMETRI POSIZIONALI E NOMINALI DI AI.AGG:\n"
                "   • Argomento 1 (posizionale): espressione riga-per-riga da aggregare, in questo caso una tupla con batch_id, affected_component e review_text.\n"
                "   • instructions => '...': parametro nominale obbligatorio che definisce il task di sintesi applicato una sola volta per l'intero gruppo.\n"
                "   • connection_id, endpoint: parametri di connessione Vertex AI e modello ('gemini-2.5-flash').\n\n"
                "2. TIPO DI RITORNO E CLAUSOLA GROUP BY:\n"
                "   • AI.AGG è una funzione aggregata SQL nativa: restituisce direttamente uno scalare STRING per gruppo (non usa .result!).\n"
                "   • Si combina nella stessa SELECT con funzioni di aggregazione tradizionali come COUNT(*) (INT64) e AVG(i.severity) (FLOAT64) sotto GROUP BY p.sku, p.product_name."
            ),
        },
    },
    {
        "id": "M8",
        "num": "08",
        "script": "sql/07_key_drivers_and_ai_agg_enhancements.sql (Query 7c + 7d)",
        "name": "Correlazione Semantica Senza Chiavi Esterne",
        "subtitle": "JOIN ON AI.IF (Correlazione Logica) & AI.SIMILARITY (Ricerca Vettoriale)",
        "goal": {
            "headline": "Collegare i bollettini tecnici dell'ingegneria con i reclami dei clienti scritti in linguaggio comune",
            "cards": [
                {
                    "color": "blue",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "Due Linguaggi che Non Si Parlano",
                    "bullets": [
                        "Il team Engineering documenta i difetti noti in bollettini tecnici con terminologia specialistica",
                        "I clienti descrivono gli stessi guasti usando parole quotidiane basate su ciò che vedono o sentono",
                        "Non esiste alcun codice o riferimento comune che colleghi una recensione cliente a un bollettino tecnico",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Collegare Sintomo Cliente e Nota Tecnica",
                    "bullets": [
                        "Quante recensioni dei clienti corrispondono esattamente alla modalità di guasto descritta in un bollettino?",
                        "Quali altre segnalazioni nel catalogo descrivono un sintomo simile a un caso critico appena scoperto?",
                        "Quando serve verificare un nesso logico causa-effetto e quando basta cercare frasi simili?",
                    ],
                },
                {
                    "color": "green",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Ponte tra Assistenza Clienti e R&D",
                    "bullets": [
                        "Misurare subito l'impatto sul campo di ciascun bollettino tecnico contando i reclami corrispondenti",
                        "Individuare rapidamente tutti i casi analoghi a una segnalazione critica descritta a parole",
                        "Unificare il patrimonio informativo dell'ingegneria con la voce del cliente senza codifiche manuali",
                    ],
                },
            ],
            "kpi": "Ponte tra Lessico Tecnico e Linguaggio Cliente  •  Impatto sul Campo dei Bollettini  •  Ricerca Casi Simili",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 08 (OBIETTIVO DI BUSINESS):\n"
                "• Mostrare come collegare due mondi che usano lessici diversi: il linguaggio ingegneristico interno (TSB) e il linguaggio colloquiale del consumatore."
            ),
        },
        "er": {
            "headline": "Join semantica logica ON AI.IF(...) e ricerca vettoriale continua con AI.SIMILARITY",
            "postit": {
                "title": "📌 JOIN AI.IF & SIMILARITY",
                "lines": [
                    "• Pre-filtro relazionale: target_sku = sku",
                    "  AND rating <= 2 (pota coppie O(N×M))",
                    "• Query 7c: INNER JOIN ON AI.IF(...)",
                    "  valuta equivalenza logica causa-effetto",
                    "• Query 7d: AI.SIMILARITY(frase, testo)",
                    "  calcola cosine similarity con embedding",
                ],
            },
            "note": "",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 08 (DIAGRAMMA E/R):\n"
                "• La relazione fisica tra engineering_bulletins e product_reviews avviene su target_sku = sku (1:N).\n"
                "• Sopra la pre-condizione relazionale equijoin (target_sku = sku AND rating <= 2) applichiamo la correlazione semantica: AI.IF logica causa-effetto in Query 7c e AI.SIMILARITY con embedding in Query 7d."
            ),
        },
        "tech": {
            "headline": "Scegliere tra verifica logica LLM (AI.IF) ed embedding vettoriali (AI.SIMILARITY)",
            "cards": [
                {
                    "color": "blue",
                    "tag": "QUERY 7c: JOIN ON AI.IF",
                    "title": "Join Relazionale + Semantica",
                    "bullets": [
                        "Usa AI.IF direttamente dentro la clausola ON di una INNER JOIN",
                        "L'LLM ragiona sul rapporto causa-effetto (es. 'deriva a 108°C' = 'non si spegne quando bolle')",
                        "Fondamentale: aggiungere sempre un pre-filtro equijoin (b.target_sku = r.sku) per evitare O(N×M)",
                    ],
                },
                {
                    "color": "green",
                    "tag": "QUERY 7d: AI.SIMILARITY",
                    "title": "Embedding Istantanei (text-embedding-005)",
                    "bullets": [
                        "Calcola al volo gli embedding dei due testi e restituisce la similarità coseno (-1..1)",
                        "Non richiede di creare una tabella di embedding né di gestire VECTOR_SEARCH",
                        "Molto più veloce ed economico di un LLM generativo per ricerche di vicinanza semantica",
                    ],
                },
                {
                    "color": "yellow",
                    "tag": "QUANDO USARE QUALE?",
                    "title": "Regola Pratica per gli Studenti",
                    "bullets": [
                        "Usa AI.SIMILARITY per trovare testi 'simili' o duplicati e ordinarli per punteggio",
                        "Usa JOIN ON AI.IF quando serve un criterio logico complesso ('A è causato da B')",
                        "Con optimization_mode => 'MINIMIZE_COST', AI.IF diventa sostenibile anche su join ampie",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 08 (PARTICOLARITÀ TECNICHE):\n"
                "• Avvertire sempre gli studenti sul costo computazionale di una JOIN semantica pura: senza b.target_sku = r.sku e r.rating <= 2, "
                "BigQuery dovrebbe valutare M × N coppie con l'LLM!"
            ),
        },
        "sql": {
            "headline": "Query 7c & 7d — Semantic JOIN con AI.IF e ricerca vettoriale con AI.SIMILARITY",
            "code": [
                "-- 7c. Semantic JOIN tra Bollettini Tecnici (b) e Recensioni (r):",
                "SELECT b.bulletin_id, COUNT(r.review_id) AS matching_reviews",
                "FROM engineering_bulletins b",
                "JOIN `mfg_quality_demo.product_reviews` r",
                "  ON b.target_sku = r.sku AND r.rating <= 2",
                " AND AI.IF(('Review: ', r.review_text, ' matches bulletin: ',",
                "            b.technical_description),",
                "   connection_id => 'us.vertex-ai-conn',",
                "   optimization_mode => 'MINIMIZE_COST') GROUP BY 1;",
                "-- 7d. Ricerca Semantica Continua con Embeddings:",
                "SELECT r.review_id, r.sku, r.batch_id,",
                "  ROUND(AI.SIMILARITY(",
                "    'kettle never stops boiling and smells like burnt plastic',",
                "    r.review_text, connection_id => 'us.vertex-ai-conn',",
                "    endpoint => 'text-embedding-005'), 3) AS semantic_similarity",
                "FROM `mfg_quality_demo.product_reviews` r",
                "ORDER BY semantic_similarity DESC LIMIT 10;",
            ],
            "callouts": [
                {
                    "color": "blue",
                    "title": "🔵 JOIN ... ON ... AND AI.IF(prompt, ...) ➔ BOOL",
                    "text": "In (7c): tupla (r.review_text, b.technical_description), connection_id, optimization_mode. Out: BOOL nella clausola ON.",
                },
                {
                    "color": "green",
                    "title": "🟢 AI.SIMILARITY(content1, content2, ...) ➔ FLOAT64",
                    "text": "In (7d): 2 argomenti posizionali STRING (testo cercato, r.review_text), connection_id, endpoint => 'text-embedding-005'.",
                },
                {
                    "color": "yellow",
                    "title": "🟡 Output 7c (INT64) e 7d (FLOAT64)",
                    "text": "Out (7c): COUNT(r.review_id) (INT64) per bulletin_id. Out (7d): scalare FLOAT64 diretto (-1.0..1.0) in ORDER BY.",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 08 (LISTATO SQL):\n\n"
                "1. PARAMETRI E SINTASSI QUERY 7c (SEMANTIC JOIN CON AI.IF):\n"
                "   • Input: tupla (r.review_text, b.technical_description), connection_id, optimization_mode => 'MINIMIZE_COST'.\n"
                "   • Output: valore BOOL valutato dinamicamente nella clausola ON della INNER JOIN insieme al pre-filtro equijoin b.target_sku = r.sku.\n\n"
                "2. PARAMETRI E SINTASSI QUERY 7d (RICERCA CON AI.SIMILARITY):\n"
                "   • Input: due stringhe posizionali (query testuale e colonna r.review_text), connection_id, endpoint => 'text-embedding-005'.\n"
                "   • Output: scalare FLOAT64 (-1.0..1.0) rappresentante la similarità coseno, pronto per ORDER BY semantic_similarity DESC LIMIT 10."
            ),
        },
    },
    {
        "id": "M9",
        "num": "09",
        "script": "sql/07_key_drivers_and_ai_agg_enhancements.sql (Query 7e + 7f + 7g)",
        "name": "Analisi Zero-Shot delle Serie Storiche",
        "subtitle": "TimesFM 2.5: AI.DETECT_ANOMALIES, AI.FORECAST Multivariato & AI.EVALUATE",
        "goal": {
            "headline": "Rilevare tempestivamente i picchi anomali di difetti e prevedere il carico futuro di reclami",
            "cards": [
                {
                    "color": "red",
                    "tag": "IL PROBLEMA DI BUSINESS",
                    "title": "Allerta Precoce & Pianificazione Scorte",
                    "bullets": [
                        "Un'impennata improvvisa di reclami dopo una spedizione deve far scattare subito un'allerta qualità",
                        "Il team Ricambi e il Customer Care devono sapere in anticipo quanti reclami arriveranno la prossima settimana",
                        "I team operativi non hanno il tempo di progettare e mantenere modelli statistici dedicati per ogni metrica",
                    ],
                },
                {
                    "color": "blue",
                    "tag": "LE DOMANDE DI BUSINESS",
                    "title": "Anomalie, Previsioni e Affidabilità",
                    "bullets": [
                        "In quali giornate il numero di difetti segnalati ha superato in modo anomalo l'andamento storico normale?",
                        "Quanti difetti e di quale gravità media dobbiamo aspettarci nei prossimi 7 giorni?",
                        "Quanto sono affidabili e accurate queste previsioni se confrontate con i dati reali osservati?",
                    ],
                },
                {
                    "color": "green",
                    "tag": "OBIETTIVO ATTESO",
                    "title": "Monitoraggio Proattivo della Qualità",
                    "bullets": [
                        "Identificare automaticamente i giorni di picco anomalo non appena i difetti deviano dal trend atteso",
                        "Prevedere congiuntamente volume dei difetti e gravità media sulla settimana successiva",
                        "Misurare oggettivamente il margine di errore previsionale per pianificare scorte e assistenza",
                    ],
                },
            ],
            "kpi": "Rilevamento Automatico dei Picchi Anomali  •  Previsione Congiunta a 7 Giorni  •  Misura dell'Accuratezza",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 09 (OBIETTIVO DI BUSINESS):\n"
                "• Chiudiamo il workshop mostrando come monitorare l'evoluzione nel tempo dei difetti e prevedere i volumi futuri con TimesFM 2.5."
            ),
        },
        "er": {
            "headline": "Aggregazione giornaliera dei difetti in una serie storica per TimesFM 2.5",
            "postit": {
                "title": "📌 Suite TimesFM 2.5",
                "lines": [
                    "• Aggregazione: GROUP BY review_date",
                    "• 7e. AI.DETECT_ANOMALIES:",
                    "  confronta storico (<10/03) vs target",
                    "• 7f. AI.FORECAST multivariato:",
                    "  target [defect_count, avg_severity]",
                    "• 7g. AI.EVALUATE: MAE, RMSE, sMAPE",
                ],
            },
            "note": "",
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 09 (DIAGRAMMA E/R):\n"
                "• Il diagramma mostra la join 1:1 su review_id tra product_reviews e review_insights e la proiezione SELECT verso la serie storica.\n"
                "• Le 3 TVF zero-shot di TimesFM 2.5: AI.DETECT_ANOMALIES (confronto storico vs target), AI.FORECAST multivariato (previsione congiunta volumi e gravità) e AI.EVALUATE (metriche MAE, RMSE, sMAPE)."
            ),
        },
        "tech": {
            "headline": "Funzioni foundation zero-shot per serie storiche integrate direttamente in BigQuery SQL",
            "cards": [
                {
                    "color": "red",
                    "tag": "7e. AI.DETECT_ANOMALIES",
                    "title": "Confronto Storico vs Finestra Target",
                    "bullets": [
                        "Prende due TABLE(): lo storico baseline (< 2025-03-10) e la finestra da ispezionare (>= 2025-03-10)",
                        "anomaly_prob_threshold => 0.80 imposta la sensibilità degli intervalli di confidenza",
                        "Restituisce is_anomaly (BOOL), lower_bound, upper_bound e anomaly_probability",
                    ],
                },
                {
                    "color": "blue",
                    "tag": "7f. AI.FORECAST MULTIVARIATO",
                    "title": "Multi-Target + Covariate Storiche",
                    "bullets": [
                        "target_cols => ['defect_count', 'avg_severity'] prevede più metriche simultaneamente",
                        "past_covariate_cols => ['total_reviews'] usa il volume totale recensioni per affinare la stima",
                        "horizon => 7 proietta i 7 giorni successivi con intervalli di confidenza al 90%",
                    ],
                },
                {
                    "color": "green",
                    "tag": "7g. AI.EVALUATE",
                    "title": "Backtesting Zero-Shot Immediato",
                    "bullets": [
                        "Confronta la finestra di training con i dati reali osservati nella finestra di holdout",
                        "Calcola istantaneamente mean_absolute_error (MAE), root_mean_squared_error (RMSE) e sMAPE",
                        "Permette di validare l'affidabilità di TimesFM 2.5 prima di metterlo in produzione",
                    ],
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 09 (PARTICOLARITÀ TECNICHE):\n"
                "• Sottolineare che nessuna di queste 3 funzioni richiede CREATE MODEL né connection_id: TimesFM 2.5 è gestito nativamente da BigQuery."
            ),
        },
        "sql": {
            "headline": "Query 7e, 7f & 7g — Rilevamento anomalie, previsione multivariata e backtesting TimesFM",
            "code": [
                "-- Condivisa: WITH daily_defects AS (SELECT TIMESTAMP(review_date)...",
                "--   COUNTIF(is_defect_report) AS defect_count, AVG(severity)...)",
                "-- 7e. Rilevamento Anomalie Zero-Shot:",
                "SELECT * FROM AI.DETECT_ANOMALIES(",
                "  TABLE(SELECT * FROM daily_defects WHERE review_date <  '2025-03-10'),",
                "  TABLE(SELECT * FROM daily_defects WHERE review_date >= '2025-03-10'),",
                "  data_col => 'defect_count', timestamp_col => 'review_date',",
                "  model => 'TimesFM 2.5', anomaly_prob_threshold => 0.80);",
                "-- 7f. Previsione Multivariata con Covariate Storiche:",
                "SELECT * FROM AI.FORECAST(TABLE(daily_defects),",
                "  timestamp_col => 'review_date',",
                "  target_cols => ['defect_count', 'avg_severity'],",
                "  past_covariate_cols => ['total_reviews'],",
                "  model => 'TimesFM 2.5', horizon => 7, confidence_level => 0.90);",
                "-- 7g. Valutazione Backtesting Zero-Shot (MAE, RMSE, sMAPE):",
                "SELECT * FROM AI.EVALUATE(",
                "  TABLE(SELECT * FROM daily_defects WHERE review_date <  '2025-03-10'),",
                "  TABLE(SELECT * FROM daily_defects WHERE review_date >= '2025-03-10'),",
                "  data_col => 'defect_count', timestamp_col => 'review_date',",
                "  model => 'TimesFM 2.5');",
            ],
            "callouts": [
                {
                    "color": "red",
                    "title": "🔴 AI.DETECT_ANOMALIES(TABLE(history), TABLE(target), ...)",
                    "text": "In (7e): 2 TABLE(...), data_col, timestamp_col, model, anomaly_prob_threshold (FLOAT64). Out: tabella con is_anomaly (BOOL).",
                },
                {
                    "color": "blue",
                    "title": "🔵 AI.FORECAST(TABLE(...), target_cols => [...], ...)",
                    "text": "In (7f): TABLE, timestamp_col, target_cols (ARRAY<STRING>), past_covariate_cols (ARRAY<STRING>), horizon (INT64), confidence_level.",
                },
                {
                    "color": "green",
                    "title": "🟢 AI.EVALUATE(TABLE(history), TABLE(actuals), ...)",
                    "text": "In (7g): 2 TABLE(...), data_col, timestamp_col, model => 'TimesFM 2.5'. Out: tabella con MAE, RMSE e sMAPE (FLOAT64).",
                },
            ],
            "notes": (
                "NOTE PER IL DOCENTE — MODULO 09 (LISTATO SQL):\n\n"
                "1. PARAMETRI DI AI.DETECT_ANOMALIES (7e):\n"
                "   • Input: due TABLE() (serie storica di riferimento e finestra target da ispezionare), data_col => 'defect_count', timestamp_col => 'review_date', model => 'TimesFM 2.5' e anomaly_prob_threshold => 0.90.\n"
                "   • Output: tabella contenente la colonna booleana is_anomaly (BOOL) e i limiti di confidenza inferiore e superiore.\n\n"
                "2. PARAMETRI DI AI.FORECAST (7f) E AI.EVALUATE (7g):\n"
                "   • AI.FORECAST: prende TABLE(storico), timestamp_col, target_cols multivariati ['defect_count', 'avg_severity'], past_covariate_cols ['total_reviews'] e horizon => 7. Restituisce le stime puntuali per i 7 giorni successivi.\n"
                "   • AI.EVALUATE: confronta due TABLE() (storico e valori effettivi) e produce 1 riga con le metriche di errore MAE, RMSE e sMAPE (FLOAT64)."
            ),
        },
    },
]

# ==============================================================================
# SQL SYNTAX SEMANTIC HIGHLIGHTER (FOR HTML)
# ==============================================================================
AI_FUNCS = [
    "AI.IF",
    "AI.CLASSIFY",
    "AI.GENERATE_INT",
    "AI.SCORE",
    "AI.GENERATE",
    "AI.KEY_DRIVERS",
    "AI.AGG",
    "AI.SIMILARITY",
    "AI.DETECT_ANOMALIES",
    "AI.FORECAST",
    "AI.EVALUATE",
]

GREEN_TOKENS = [
    "output_schema",
    "optimization_mode",
    "'MINIMIZE_COST'",
    "categories",
    "instructions",
    "target_cols",
    "past_covariate_cols",
    "anomaly_prob_threshold",
    "'TimesFM 2.5'",
    "'gemini-2.5-flash'",
    "'text-embedding-005'",
    "affected_component",
    "component_type",
]

RED_TOKENS = [
    "REPEAT",
    "UNTIL",
    "END REPEAT",
    "DELETE FROM",
    "retryable error",
    "safety_hazard_score",
    "installed_by_machine_id",
    "last_maintenance_date",
    "is_implicated_defect",
]

YELLOW_TOKENS = [
    "LEFT JOIN",
    "JOIN",
    "IS NULL",
    "GROUP BY",
    "STRING_AGG",
    "SAFE_DIVIDE",
    "COUNTIF",
    "TABLE",
]


def highlight_sql_html(lines):
    out_lines = []
    for line in lines:
        esc = html.escape(line)
        if esc.strip().startswith("--"):
            out_lines.append(f'<span class="tok-comment">{esc}</span>')
            continue
        for t in AI_FUNCS:
            esc = re.sub(
                rf"\b({re.escape(t)})\b",
                r'<span class="tok-blue">\1</span>',
                esc,
            )
        for t in GREEN_TOKENS:
            esc_t = html.escape(t)
            esc = esc.replace(esc_t, f'<span class="tok-green">{esc_t}</span>')
        for t in RED_TOKENS:
            esc_t = html.escape(t)
            esc = esc.replace(esc_t, f'<span class="tok-red">{esc_t}</span>')
        for t in YELLOW_TOKENS:
            esc_t = html.escape(t)
            esc = re.sub(
                rf"\b({re.escape(esc_t)})\b",
                r'<span class="tok-yellow">\1</span>',
                esc,
            )
        out_lines.append(esc)
    return "\n".join(out_lines)


def render_postit_html(postit):
    items = "".join(f"<li>{html.escape(line.lstrip('• ').strip())}</li>" for line in postit["lines"])
    return f"""
    <aside class="postit-note">
      <div class="postit-tape"></div>
      <h4>{html.escape(postit['title'])}</h4>
      <ul>{items}</ul>
    </aside>
    """


# ==============================================================================
# HTML GENERATOR (STANDALONE 38-SLIDE 16:9 DECK)
# ==============================================================================
def build_html_deck():
    slides_html = []
    er_json_path = os.path.join(DECK_DIR, "er_rendered.json")
    er_rendered = json.load(open(er_json_path, encoding="utf-8")) if os.path.exists(er_json_path) else {}

    def get_mermaid_er_html(slide_key):
        if slide_key in er_rendered:
            mmd_code = html.escape(er_rendered[slide_key]["mmd"])
            svg_markup = er_rendered[slide_key]["svg"]
            return f'<pre class="mermaid-source" style="display:none">{mmd_code}</pre>\n{svg_markup}'
        return "<p>Diagramma E/R</p>"

    # SLIDE 1: COVER
    slides_html.append(
        """
    <section class="slide cover-slide">
      <div class="rainbow-bar">
        <span style="background:#4285F4"></span>
        <span style="background:#EA4335"></span>
        <span style="background:#FBBC04"></span>
        <span style="background:#34A853"></span>
      </div>
      <div class="cover-top">
        <span class="badge badge-blue">WORKSHOP GOOGLE CLOUD APPLIED AI • LAB II (PRODUCT_ANALYTICS)</span>
        <span class="conf-label">Slide Deck per il Docente • Font ≥ 10pt</span>
      </div>
      <div class="cover-main">
        <h1>NovaHome — Analisi Intelligente delle Recensioni</h1>
        <h2>Dal Testo Non Strutturato alla Causa Radice di Fabbrica con <span style="color:#4285F4">BigQuery AI</span></h2>
        <p class="cover-desc">Guida visuale alle <strong>9 Query Avanzate</strong> del laboratorio: Obiettivo di Business, Diagramma Concettuale E/R (Mermaid <code>erDiagram</code>), Listato SQL Sintetico a Colori Semantici e Particolarità Tecniche.</p>
      </div>
      <div class="semantic-legend">
        <div class="leg-item leg-blue"><strong>🔵 BLU — Dati Grezzi &amp; Funzioni AI</strong><span>product_reviews • AI.IF, AI.CLASSIFY, AI.SCORE, AI.GENERATE, AI.AGG, AI.FORECAST</span></div>
        <div class="leg-item leg-green"><strong>🟢 VERDE — Tabelle AI &amp; Schema Tipizzato</strong><span>review_insights • output_schema • MINIMIZE_COST • TimesFM 2.5</span></div>
        <div class="leg-item leg-yellow"><strong>🟡 GIALLO — Distinta Base (BOM) &amp; Join SQL</strong><span>products, batches, batch_components, component_lots • LEFT JOIN, GROUP BY</span></div>
        <div class="leg-item leg-red"><strong>🔴 ROSSO — Causa Radice &amp; Resilienza</strong><span>machines (CAL-02) • Lotto HE-4471 (77%) • Ciclo REPEAT/UNTIL 429</span></div>
      </div>
      <div class="slide-footer"><span>Google Cloud • product_analytics</span><span>Slide 1 / 38</span></div>
    </section>
    """
    )

    # SLIDE 2: GLOBAL E/R (Full-width diagram, no sticky note, no bottom banner)
    slides_html.append(
        f"""
    <section class="slide">
      <div class="slide-header">
        <div>
          <span class="badge badge-yellow">ARCHITETTURA DATI GLOBALE • MERMAID erDiagram • DATASET mfg_quality_demo</span>
          <h2>Mappa E/R Completa</h2>
        </div>
        <span class="script-ref">sql/01_setup_dataset_and_data.sql</span>
      </div>
      <div class="er-container" style="flex:1; margin: 10px 0;">
        {get_mermaid_er_html("slide_02")}
      </div>
      <div class="slide-footer"><span>Mappa E/R Completa</span><span>Slide 2 / 38</span></div>
    </section>
    """
    )

    slide_num = 3
    for m in MODULES:
        # SLIDE .1: OBIETTIVO DI BUSINESS
        cards_html = ""
        for c in m["goal"]["cards"]:
            bullets = "".join(f"<li>{html.escape(b)}</li>" for b in c["bullets"])
            cards_html += f"""
            <div class="card card-{c['color']}">
              <div class="card-tag">{html.escape(c['tag'])}</div>
              <h3>{html.escape(c['title'])}</h3>
              <ul>{bullets}</ul>
            </div>
            """
        slides_html.append(
            f"""
        <section class="slide" data-notes="{html.escape(m['goal']['notes'])}">
          <div class="slide-header">
            <div>
              <span class="badge badge-blue">MODULO {m['num']} • OBIETTIVO DI BUSINESS • {html.escape(m['name'].upper())}</span>
              <h2>{html.escape(m['goal']['headline'])}</h2>
            </div>
            <span class="script-ref">{html.escape(m['script'])}</span>
          </div>
          <div class="cards-grid-3">
            {cards_html}
          </div>
          <div class="bottom-banner banner-blue">
            <strong>🎯 Punto Chiave:</strong> {html.escape(m['goal']['kpi'])}
          </div>
          <div class="slide-footer"><span>Modulo {m['num']}: {html.escape(m['name'])}</span><span>Slide {slide_num} / 38</span></div>
        </section>
        """
        )
        slide_num += 1

        # SLIDE .2: DIAGRAMMA CONCETTUALE E/R (MERMAID erDiagram)
        er_slide_key = f"slide_{slide_num:02d}"
        slides_html.append(
            f"""
        <section class="slide" data-notes="{html.escape(m['er']['notes'])}">
          <div class="slide-header">
            <div>
              <span class="badge badge-yellow">MODULO {m['num']} • DIAGRAMMA E/R • {html.escape(m['name'].upper())}</span>
              <h2>{html.escape(m['er']['headline'])}</h2>
            </div>
            <span class="script-ref">{html.escape(m['script'])}</span>
          </div>
          <div class="er-container" style="flex:1; margin: 10px 0;">
            {get_mermaid_er_html(er_slide_key)}
          </div>
          <div class="slide-footer"><span>Modulo {m['num']}: {html.escape(m['name'])}</span><span>Slide {slide_num} / 38</span></div>
        </section>
        """
        )
        slide_num += 1

        # SLIDE .3: LISTATO SQL SINTETICO (ora subito dopo il Diagramma E/R!)
        callouts_html = ""
        for cal in m["sql"]["callouts"]:
            callouts_html += f"""
            <div class="sql-callout callout-{cal['color']}">
              <h4>{html.escape(cal['title'])}</h4>
              <p>{html.escape(cal['text'])}</p>
            </div>
            """
        sql_colored = highlight_sql_html(m["sql"]["code"])
        slides_html.append(
            f"""
        <section class="slide" data-notes="{html.escape(m['sql']['notes'])}">
          <div class="slide-header">
            <div>
              <span class="badge badge-red">MODULO {m['num']} • LISTATO SQL SINTETICO • {html.escape(m['name'].upper())}</span>
              <h2>{html.escape(m['sql']['headline'])}</h2>
            </div>
            <span class="script-ref">{html.escape(m['script'])}</span>
          </div>
          <div class="sql-layout">
            <pre class="sql-box"><code>{sql_colored}</code></pre>
            <div class="sql-sidebar">
              {callouts_html}
            </div>
          </div>
          <div class="slide-footer"><span>Modulo {m['num']}: {html.escape(m['name'])}</span><span>Slide {slide_num} / 38</span></div>
        </section>
        """
        )
        slide_num += 1

        # SLIDE .4: PARTICOLARITÀ TECNICHE
        tcards_html = ""
        for c in m["tech"]["cards"]:
            bullets = "".join(f"<li>{html.escape(b)}</li>" for b in c["bullets"])
            tcards_html += f"""
            <div class="card card-{c['color']}">
              <div class="card-tag">{html.escape(c['tag'])}</div>
              <h3>{html.escape(c['title'])}</h3>
              <ul>{bullets}</ul>
            </div>
            """
        slides_html.append(
            f"""
        <section class="slide" data-notes="{html.escape(m['tech']['notes'])}">
          <div class="slide-header">
            <div>
              <span class="badge badge-green">MODULO {m['num']} • PARTICOLARITÀ TECNICHE • {html.escape(m['name'].upper())}</span>
              <h2>{html.escape(m['tech']['headline'])}</h2>
            </div>
            <span class="script-ref">{html.escape(m['script'])}</span>
          </div>
          <div class="cards-grid-3">
            {tcards_html}
          </div>
          <div class="bottom-banner banner-green">
            <strong>⚙️ Costrutto Chiave:</strong> {html.escape(m['subtitle'])}
          </div>
          <div class="slide-footer"><span>Modulo {m['num']}: {html.escape(m['name'])}</span><span>Slide {slide_num} / 38</span></div>
        </section>
        """
        )
        slide_num += 1

    html_doc = f"""<!DOCTYPE html>
<html lang="it">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>NovaHome — Analisi Intelligente delle Recensioni (product_analytics)</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto+Mono:wght@400;600;700&family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
  <style>
    :root {{
      --slide-w: 960px;
      --slide-h: 540px;
      --dark: #202124;
      --grey: #5F6368;
      --blue: #1967D2;
      --blue-bg: #E8F0FE;
      --blue-bd: #8AB4F8;
      --green: #137333;
      --green-bg: #E6F4EA;
      --green-bd: #81C995;
      --yellow: #B06000;
      --yellow-bg: #FEF7E0;
      --yellow-bd: #FDD663;
      --red: #C5221F;
      --red-bg: #FCE8E6;
      --red-bd: #F28B82;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      background: #E8EAED;
      font-family: 'Google Sans', 'Roboto', sans-serif;
      color: var(--dark);
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 24px 0 60px;
      gap: 28px;
    }}
    .deck-toolbar {{
      position: sticky;
      top: 12px;
      z-index: 100;
      background: rgba(32, 33, 36, 0.92);
      color: #fff;
      padding: 8px 18px;
      border-radius: 999px;
      font-size: 13px;
      display: flex;
      align-items: center;
      gap: 16px;
      box-shadow: 0 4px 14px rgba(0,0,0,0.2);
    }}
    .deck-toolbar button {{
      background: #4285F4;
      color: #fff;
      border: none;
      padding: 5px 12px;
      border-radius: 999px;
      cursor: pointer;
      font-weight: 600;
      font-size: 12px;
    }}
    .slide {{
      width: var(--slide-w);
      height: var(--slide-h);
      background: #FFFFFF;
      border-radius: 8px;
      box-shadow: 0 6px 20px rgba(0,0,0,0.10);
      padding: 28px 36px 22px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      position: relative;
      overflow: hidden;
      page-break-after: always;
      break-after: page;
    }}
    .rainbow-bar {{
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 6px;
      display: grid;
      grid-template-columns: repeat(4, 1fr);
    }}
    .cover-top {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 8px;
    }}
    .conf-label {{ font-size: 12px; color: var(--grey); font-weight: 500; }}
    .cover-main h1 {{ font-size: 38px; font-weight: 700; color: var(--dark); line-height: 1.15; margin-bottom: 10px; }}
    .cover-main h2 {{ font-size: 22px; font-weight: 500; color: var(--grey); margin-bottom: 14px; }}
    .cover-desc {{ font-size: 14.5px; color: var(--dark); max-width: 820px; line-height: 1.45; }}
    .semantic-legend {{
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 12px;
      margin-top: 8px;
    }}
    .leg-item {{
      padding: 10px 14px;
      border-radius: 6px;
      border-left: 5px solid;
      display: flex;
      flex-direction: column;
      gap: 3px;
    }}
    .leg-item strong {{ font-size: 13px; }}
    .leg-item span {{ font-size: 12px; color: var(--dark); font-family: 'Roboto Mono', monospace; }}
    .leg-blue {{ background: var(--blue-bg); border-color: var(--blue); color: var(--blue); }}
    .leg-green {{ background: var(--green-bg); border-color: var(--green); color: var(--green); }}
    .leg-yellow {{ background: var(--yellow-bg); border-color: var(--yellow); color: var(--yellow); }}
    .leg-red {{ background: var(--red-bg); border-color: var(--red); color: var(--red); }}

    .slide-header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 16px;
      border-bottom: 1px solid #E8EAED;
      padding-bottom: 10px;
    }}
    .slide-header h2 {{
      font-size: 20.5px;
      font-weight: 600;
      color: var(--dark);
      margin-top: 5px;
      line-height: 1.22;
    }}
    .script-ref {{
      font-family: 'Roboto Mono', monospace;
      font-size: 11.5px;
      background: #F1F3F4;
      color: var(--grey);
      padding: 4px 10px;
      border-radius: 4px;
      white-space: nowrap;
      flex-shrink: 0;
    }}
    .badge {{
      display: inline-block;
      font-size: 11px;
      font-weight: 700;
      padding: 3px 9px;
      border-radius: 4px;
      letter-spacing: 0.4px;
    }}
    .badge-blue {{ background: var(--blue-bg); color: var(--blue); }}
    .badge-green {{ background: var(--green-bg); color: var(--green); }}
    .badge-yellow {{ background: var(--yellow-bg); color: var(--yellow); }}
    .badge-red {{ background: var(--red-bg); color: var(--red); }}

    .cards-grid-3 {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 16px;
      flex: 1;
      margin: 14px 0;
    }}
    .card {{
      border-radius: 8px;
      padding: 16px 16px 14px;
      border: 1.5px solid #DADCE0;
      border-top-width: 6px;
      background: #FFFFFF;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }}
    .card-blue {{ border-top-color: #4285F4; background: #FAFCFF; }}
    .card-green {{ border-top-color: #34A853; background: #F9FDFB; }}
    .card-yellow {{ border-top-color: #FBBC04; background: #FFFDF7; }}
    .card-red {{ border-top-color: #EA4335; background: #FFF9F9; }}

    .card-tag {{
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.4px;
    }}
    .card-blue .card-tag {{ color: var(--blue); }}
    .card-green .card-tag {{ color: var(--green); }}
    .card-yellow .card-tag {{ color: var(--yellow); }}
    .card-red .card-tag {{ color: var(--red); }}

    .card h3 {{
      font-size: 16.5px;
      font-weight: 700;
      color: var(--dark);
      line-height: 1.25;
    }}
    .card ul {{
      padding-left: 18px;
      display: flex;
      flex-direction: column;
      gap: 8px;
      margin-top: 4px;
    }}
    .card li {{
      font-size: 13.5px;
      line-height: 1.38;
      color: #3C4043;
    }}

    .er-split-layout {{
      flex: 1;
      display: grid;
      grid-template-columns: 71% 27%;
      gap: 2%;
      align-items: center;
      margin: 10px 0;
      min-height: 0;
    }}
    .er-container {{
      height: 100%;
      display: flex;
      align-items: center;
      justify-content: center;
      background: #F8F9FA;
      border: 1px solid #E8EAED;
      border-radius: 8px;
      padding: 6px 10px;
      overflow: hidden;
    }}
    .er-container svg {{
      max-width: 100%;
      max-height: 315px;
      width: auto;
      height: auto;
    }}
    .postit-note {{
      background: #FFF9C4;
      border: 1px solid #FBC02D;
      border-radius: 4px;
      padding: 16px 14px 14px;
      box-shadow: 3px 5px 12px rgba(0,0,0,0.12);
      position: relative;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }}
    .postit-tape {{
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 8px;
      background: #FBC02D;
      border-radius: 3px 3px 0 0;
    }}
    .postit-note h4 {{
      font-size: 13.5px;
      font-weight: 700;
      color: #5D4037;
      margin-top: 2px;
    }}
    .postit-note ul {{
      padding-left: 16px;
      display: flex;
      flex-direction: column;
      gap: 5px;
    }}
    .postit-note li {{
      font-size: 12px;
      line-height: 1.35;
      color: #202124;
    }}

    .sql-layout {{
      display: grid;
      grid-template-columns: 64% 34%;
      gap: 2%;
      flex: 1;
      margin: 12px 0;
      align-items: stretch;
    }}
    .sql-box {{
      background: #F8F9FA;
      border: 1.5px solid #DADCE0;
      border-left: 5px solid #4285F4;
      border-radius: 6px;
      padding: 12px 14px;
      font-family: 'Roboto Mono', monospace;
      font-size: 12.2px;
      line-height: 1.38;
      color: #202124;
      overflow: hidden;
      white-space: pre;
      display: flex;
      align-items: center;
    }}
    .tok-blue {{ color: #1967D2; font-weight: 700; background: #E8F0FE; padding: 0 3px; border-radius: 3px; }}
    .tok-green {{ color: #137333; font-weight: 700; background: #E6F4EA; padding: 0 3px; border-radius: 3px; }}
    .tok-yellow {{ color: #B06000; font-weight: 700; background: #FEF7E0; padding: 0 3px; border-radius: 3px; }}
    .tok-red {{ color: #C5221F; font-weight: 700; background: #FCE8E6; padding: 0 3px; border-radius: 3px; }}
    .tok-comment {{ color: #5F6368; font-style: italic; }}

    .sql-sidebar {{
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      gap: 10px;
    }}
    .sql-callout {{
      padding: 12px 14px;
      border-radius: 6px;
      border-left: 5px solid;
      flex: 1;
      display: flex;
      flex-direction: column;
      justify-content: center;
      gap: 4px;
    }}
    .sql-callout h4 {{ font-size: 13.5px; font-weight: 700; }}
    .sql-callout p {{ font-size: 12.5px; line-height: 1.35; color: var(--dark); }}
    .callout-blue {{ background: var(--blue-bg); border-color: var(--blue); color: var(--blue); }}
    .callout-green {{ background: var(--green-bg); border-color: var(--green); color: var(--green); }}
    .callout-yellow {{ background: var(--yellow-bg); border-color: var(--yellow); color: var(--yellow); }}
    .callout-red {{ background: var(--red-bg); border-color: var(--red); color: var(--red); }}

    .bottom-banner {{
      padding: 9px 14px;
      border-radius: 6px;
      font-size: 13px;
      line-height: 1.35;
      border-left: 5px solid;
    }}
    .banner-blue {{ background: var(--blue-bg); border-color: var(--blue); color: var(--dark); }}
    .banner-green {{ background: var(--green-bg); border-color: var(--green); color: var(--dark); }}
    .banner-yellow {{ background: var(--yellow-bg); border-color: var(--yellow); color: var(--dark); }}
    .banner-red {{ background: var(--red-bg); border-color: var(--red); color: var(--dark); }}

    .slide-footer {{
      display: flex;
      justify-content: space-between;
      font-size: 11px;
      color: var(--grey);
      padding-top: 6px;
      border-top: 1px solid #F1F3F4;
    }}

    @media print {{
      @page {{ size: 960px 540px; margin: 0; }}
      body {{ background: #fff; padding: 0; gap: 0; }}
      .deck-toolbar {{ display: none !important; }}
      .slide {{ box-shadow: none; border-radius: 0; width: 960px; height: 540px; }}
    }}
  </style>
</head>
<body>
  <div class="deck-toolbar">
    <span><strong>NovaHome Product Analytics</strong> — 38 Slide per il Docente (9 Moduli × 4 Pagine)</span>
    <button onclick="window.print()">🖨️ Stampa / Salva PDF</button>
  </div>
  {"".join(slides_html)}
</body>
</html>
"""
    html_path = os.path.join(DECK_DIR, "slides.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_doc)
    print(f"Wrote {html_path} ({len(slides_html)} slides)")


# ==============================================================================
# GOOGLE SLIDES / PPTX BATCH JSON GENERATOR (720 x 405 PT CANVAS)
# ==============================================================================
def utf16_len(s: str) -> int:
    return len(s.encode("utf-16-le")) // 2


def build_gslides_batch():
    ops = [
        {"op": "delete-element", "element": "i0"},
        {"op": "delete-element", "element": "i1"},
    ]

    # SLIDE 1: COVER (`p`)
    ops.extend(
        [
            {"op": "set-background", "slide": "p", "color": "#FFFFFF"},
            {
                "op": "add-shape",
                "slide": "p",
                "shape_type": "RECTANGLE",
                "x": 0,
                "y": 0,
                "width": 180,
                "height": 6,
                "background_color": "#4285F4",
            },
            {
                "op": "add-shape",
                "slide": "p",
                "shape_type": "RECTANGLE",
                "x": 180,
                "y": 0,
                "width": 180,
                "height": 6,
                "background_color": "#EA4335",
            },
            {
                "op": "add-shape",
                "slide": "p",
                "shape_type": "RECTANGLE",
                "x": 360,
                "y": 0,
                "width": 180,
                "height": 6,
                "background_color": "#FBBC04",
            },
            {
                "op": "add-shape",
                "slide": "p",
                "shape_type": "RECTANGLE",
                "x": 540,
                "y": 0,
                "width": 180,
                "height": 6,
                "background_color": "#34A853",
            },
            {
                "op": "add-textbox",
                "slide": "p",
                "text": "WORKSHOP GOOGLE CLOUD APPLIED AI • LAB II (PRODUCT_ANALYTICS)",
                "x": 36,
                "y": 24,
                "width": 440,
                "height": 22,
                "font_size": 10,
                "bold": True,
                "color": "#1967D2",
                "background_color": "#E8F0FE",
                "font_family": "Roboto",
                "content_alignment": "MIDDLE",
            },
            {
                "op": "add-textbox",
                "slide": "p",
                "text": "NovaHome — Analisi Intelligente delle Recensioni",
                "x": 36,
                "y": 55,
                "width": 648,
                "height": 48,
                "font_size": 27,
                "bold": True,
                "color": "#202124",
                "font_family": "Roboto",
            },
            {
                "op": "add-textbox",
                "slide": "p",
                "text": "Dal Testo Non Strutturato alla Causa Radice di Fabbrica con BigQuery AI\n9 Moduli Didattici × 4 Pagine: Obiettivo di Business • Diagramma E/R • Listato SQL • Particolarità Tecniche",
                "x": 36,
                "y": 106,
                "width": 648,
                "height": 48,
                "font_size": 13,
                "color": "#5F6368",
                "font_family": "Roboto",
            },
        ]
    )

    legend_boxes = [
        (
            36,
            170,
            "#E8F0FE",
            "#1967D2",
            "BLU — Dati Grezzi & Funzioni BigQuery AI\nproduct_reviews • AI.IF, AI.CLASSIFY, AI.SCORE, AI.GENERATE, AI.AGG, AI.FORECAST",
        ),
        (
            366,
            170,
            "#E6F4EA",
            "#137333",
            "VERDE — Tabelle Arricchite AI & Schema Tipizzato\nreview_insights • output_schema • MINIMIZE_COST • TimesFM 2.5",
        ),
        (
            36,
            255,
            "#FEF7E0",
            "#B06000",
            "GIALLO — Grafo Supply Chain (BOM) & Join SQL\nproducts, batches, batch_components, component_lots • LEFT JOIN, GROUP BY",
        ),
        (
            366,
            255,
            "#FCE8E6",
            "#C5221F",
            "ROSSO — Causa Radice Fisica & Resilienza\nmachines (CAL-02) • Lotto HE-4471 (77%) • Ciclo REPEAT/UNTIL 429",
        ),
    ]
    for lx, ly, bg, fg, ltxt in legend_boxes:
        ops.append(
            {
                "op": "add-textbox",
                "slide": "p",
                "text": ltxt,
                "x": lx,
                "y": ly,
                "width": 318,
                "height": 72,
                "font_size": 10.5,
                "bold": True,
                "color": fg,
                "background_color": bg,
                "font_family": "Roboto",
                "content_alignment": "MIDDLE",
            }
        )
    ops.append(
        {
            "op": "add-textbox",
            "slide": "p",
            "text": "Google Cloud • NovaHome — Analisi Intelligente delle Recensioni (product_analytics)                                   Slide 1 / 38",
            "x": 36,
            "y": 376,
            "width": 648,
            "height": 18,
            "font_size": 10,
            "color": "#5F6368",
            "font_family": "Roboto",
        }
    )
    ops.append(
        {
            "op": "set-notes",
            "slide": "p",
            "text": (
                "NOTE PER IL DOCENTE — COPERTINA:\n"
                "• Benvenuti al Lab II (product_analytics) del workshop Google Cloud Applied AI.\n"
                "• Ogni query del laboratorio è presentata con 4 slide: Obiettivo di Business, Diagramma E/R (Mermaid erDiagram), Listato SQL e Particolarità Tecniche."
            ),
        }
    )

    def add_header_footer(slide_id, badge_txt, badge_color, title_txt, script_txt, footer_left, snum):
        c = COLORS[badge_color]
        ops.extend(
            [
                {"op": "set-background", "slide": slide_id, "color": "#FFFFFF"},
                {
                    "op": "add-textbox",
                    "slide": slide_id,
                    "text": badge_txt,
                    "x": 36,
                    "y": 14,
                    "width": 412,
                    "height": 18,
                    "font_size": 10,
                    "bold": True,
                    "color": c["accent"],
                    "background_color": c["fill"],
                    "font_family": "Roboto",
                    "content_alignment": "MIDDLE",
                },
                {
                    "op": "add-textbox",
                    "slide": slide_id,
                    "text": script_txt,
                    "x": 454,
                    "y": 14,
                    "width": 230,
                    "height": 18,
                    "font_size": 10,
                    "color": "#5F6368",
                    "background_color": "#F1F3F4",
                    "font_family": "Roboto Mono",
                    "alignment": "center",
                    "content_alignment": "MIDDLE",
                },
                {
                    "op": "add-textbox",
                    "slide": slide_id,
                    "text": title_txt,
                    "x": 36,
                    "y": 36,
                    "width": 648,
                    "height": 32,
                    "font_size": 15.5,
                    "bold": True,
                    "color": "#202124",
                    "font_family": "Roboto",
                },
                {
                    "op": "add-textbox",
                    "slide": slide_id,
                    "text": f"{footer_left}                                                                 Slide {snum} / 38",
                    "x": 36,
                    "y": 380,
                    "width": 648,
                    "height": 18,
                    "font_size": 10,
                    "color": "#5F6368",
                    "font_family": "Roboto",
                },
            ]
        )

    def render_er_on_slide(slide_id, postit_spec, bottom_banner_txt=None, banner_color=None):
        """Renders the stylized sticky note on the right (x=514..684) and optional bottom banner, leaving x=24..502 open for the Mermaid erDiagram PNG."""
        has_banner = bool(bottom_banner_txt and banner_color)
        postit_h = 232 if has_banner else 268
        # Top yellow sticky-tape bar
        ops.append(
            {
                "op": "add-shape",
                "slide": slide_id,
                "shape_type": "RECTANGLE",
                "x": 514,
                "y": 82,
                "width": 170,
                "height": 8,
                "background_color": "#FBC02D",
            }
        )
        postit_text = postit_spec["title"] + "\n\n" + "\n".join(postit_spec["lines"])
        ops.append(
            {
                "op": "add-textbox",
                "slide": slide_id,
                "text": postit_text,
                "x": 514,
                "y": 90,
                "width": 170,
                "height": postit_h,
                "font_size": 10,
                "color": "#202124",
                "background_color": "#FFF9C4",
                "font_family": "Roboto",
            }
        )
        if has_banner:
            bc = COLORS[banner_color]
            ops.append(
                {
                    "op": "add-textbox",
                    "slide": slide_id,
                    "text": bottom_banner_txt,
                    "x": 36,
                    "y": 338,
                    "width": 648,
                    "height": 36,
                    "font_size": 10.5,
                    "bold": True,
                    "color": "#202124",
                    "background_color": bc["fill"],
                    "font_family": "Roboto",
                    "content_alignment": "MIDDLE",
                }
            )

    # SLIDE 2: GLOBAL E/R OVERVIEW (Full-width diagram, no sticky note, no bottom banner)
    ops.append({"op": "add-slide", "layout": "BLANK", "id": "SLIDE_2"})
    add_header_footer(
        "SLIDE_2",
        "ARCHITETTURA DATI GLOBALE • DATASET mfg_quality_demo",
        "yellow",
        "Mappa E/R Completa",
        "sql/01_setup_dataset_and_data.sql",
        "Mappa E/R Completa",
        2,
    )
    ops.append(
        {
            "op": "set-notes",
            "slide": "SLIDE_2",
            "text": (
                "NOTE PER IL DOCENTE — MAPPA E/R COMPLETA (7 TABELLE):\n"
                "1. TABELLA DI APPOGGIO PER RELAZIONE M:N (batch_components):\n"
                "   • La tabella batch_components è la tabella di appoggio (tabella ponte associativa) che serve a rappresentare la relazione molti-a-molti (M:N) "
                "tra production_batches (i lotti di prodotto finito) e component_lots (i lotti di componenti ricevuti dai fornitori).\n"
                "   • Ogni lotto di produzione monta 5 componenti diversi e ogni lotto di componenti viene impiegato in più batch: "
                "per questo la sua Primary Key è la chiave composta (batch_id, lot_id), dove entrambi gli attributi sono marcati PK,FK.\n"
                "2. Review_insights E’ IL RISULTATO DELL’ESERCIZIO\n"
                "   • La tabella review_insights (in verde) viene popolata da AI.GENERATE a partire da product_reviews.review_text e si collega a component_lots "
                "tramite il predicato semantico affected_component = component_type."
            ),
        }
    )

    snum = 3
    for m in MODULES:
        # PAGE 1: OBIETTIVO DI BUSINESS
        s1 = f"SLIDE_{snum}"
        ops.append({"op": "add-slide", "layout": "BLANK", "id": s1})
        add_header_footer(
            s1,
            f"MODULO {m['num']} • OBIETTIVO DI BUSINESS • {m['name'].upper()}",
            "blue",
            m["goal"]["headline"],
            m["script"],
            f"Modulo {m['num']}: {m['name']}",
            snum,
        )
        for idx, card in enumerate(m["goal"]["cards"]):
            cx = 36 + idx * 220
            cc = COLORS[card["color"]]
            ops.append(
                {
                    "op": "add-shape",
                    "slide": s1,
                    "shape_type": "RECTANGLE",
                    "x": cx,
                    "y": 78,
                    "width": 208,
                    "height": 6,
                    "background_color": cc["primary"],
                }
            )
            c_text = f"{card['tag']}\n{card['title']}\n\n" + "\n\n".join(
                f"• {b}" for b in card["bullets"]
            )
            ops.append(
                {
                    "op": "add-textbox",
                    "slide": s1,
                    "text": c_text,
                    "x": cx,
                    "y": 84,
                    "width": 208,
                    "height": 244,
                    "font_size": 10.5,
                    "color": "#202124",
                    "background_color": cc["fill"],
                    "font_family": "Roboto",
                }
            )
        ops.append(
            {
                "op": "add-textbox",
                "slide": s1,
                "text": f"Punto Chiave: {m['goal']['kpi']}",
                "x": 36,
                "y": 338,
                "width": 648,
                "height": 34,
                "font_size": 11,
                "bold": True,
                "color": "#1967D2",
                "background_color": "#E8F0FE",
                "font_family": "Roboto",
                "content_alignment": "MIDDLE",
            }
        )
        ops.append({"op": "set-notes", "slide": s1, "text": m["goal"]["notes"]})
        snum += 1

        # PAGE 2: DIAGRAMMA E/R
        s2 = f"SLIDE_{snum}"
        ops.append({"op": "add-slide", "layout": "BLANK", "id": s2})
        add_header_footer(
            s2,
            f"MODULO {m['num']} • DIAGRAMMA E/R • {m['name'].upper()}",
            "yellow",
            m["er"]["headline"],
            m["script"],
            f"Modulo {m['num']}: {m['name']}",
            snum,
        )
        ops.append({"op": "set-notes", "slide": s2, "text": m["er"]["notes"]})
        snum += 1

        # PAGE 3: LISTATO SQL SINTETICO (subito dopo il Diagramma E/R!)
        s3 = f"SLIDE_{snum}"
        ops.append({"op": "add-slide", "layout": "BLANK", "id": s3})
        add_header_footer(
            s3,
            f"MODULO {m['num']} • LISTATO SQL SINTETICO • {m['name'].upper()}",
            "red",
            m["sql"]["headline"],
            m["script"],
            f"Modulo {m['num']}: {m['name']}",
            snum,
        )
        code_str = "\n".join(m["sql"]["code"])
        ops.append(
            {
                "op": "add-textbox",
                "slide": s3,
                "text": code_str,
                "x": 36,
                "y": 76,
                "width": 436,
                "height": 296,
                "font_size": 10,
                "color": "#202124",
                "background_color": "#F8F9FA",
                "font_family": "Roboto Mono",
            }
        )
        for idx, cal in enumerate(m["sql"]["callouts"]):
            cy = 76 + idx * 101
            cc = COLORS[cal["color"]]
            ops.append(
                {
                    "op": "add-textbox",
                    "slide": s3,
                    "text": f"{cal['title']}\n{cal['text']}",
                    "x": 482,
                    "y": cy,
                    "width": 202,
                    "height": 93,
                    "font_size": 10,
                    "color": "#202124",
                    "background_color": cc["fill"],
                    "font_family": "Roboto",
                    "content_alignment": "MIDDLE",
                }
            )
        ops.append({"op": "set-notes", "slide": s3, "text": m["sql"]["notes"]})
        snum += 1

        # PAGE 4: PARTICOLARITÀ TECNICHE
        s4 = f"SLIDE_{snum}"
        ops.append({"op": "add-slide", "layout": "BLANK", "id": s4})
        add_header_footer(
            s4,
            f"MODULO {m['num']} • PARTICOLARITÀ TECNICHE • {m['name'].upper()}",
            "green",
            m["tech"]["headline"],
            m["script"],
            f"Modulo {m['num']}: {m['name']}",
            snum,
        )
        for idx, card in enumerate(m["tech"]["cards"]):
            cx = 36 + idx * 220
            cc = COLORS[card["color"]]
            ops.append(
                {
                    "op": "add-shape",
                    "slide": s4,
                    "shape_type": "RECTANGLE",
                    "x": cx,
                    "y": 78,
                    "width": 208,
                    "height": 6,
                    "background_color": cc["primary"],
                }
            )
            c_text = f"{card['tag']}\n{card['title']}\n\n" + "\n\n".join(
                f"• {b}" for b in card["bullets"]
            )
            ops.append(
                {
                    "op": "add-textbox",
                    "slide": s4,
                    "text": c_text,
                    "x": cx,
                    "y": 84,
                    "width": 208,
                    "height": 244,
                    "font_size": 10.5,
                    "color": "#202124",
                    "background_color": cc["fill"],
                    "font_family": "Roboto",
                }
            )
        ops.append(
            {
                "op": "add-textbox",
                "slide": s4,
                "text": f"Costrutto Chiave: {m['subtitle']}",
                "x": 36,
                "y": 338,
                "width": 648,
                "height": 34,
                "font_size": 11,
                "bold": True,
                "color": "#137333",
                "background_color": "#E6F4EA",
                "font_family": "Roboto",
                "content_alignment": "MIDDLE",
            }
        )
        ops.append({"op": "set-notes", "slide": s4, "text": m["tech"]["notes"]})
        snum += 1

    batch_path = os.path.join(DECK_DIR, "gslides_batch.json")
    with open(batch_path, "w", encoding="utf-8") as f:
        json.dump(ops, f, indent=2, ensure_ascii=False)
    print(f"Wrote {batch_path} ({len(ops)} batch operations)")


if __name__ == "__main__":
    build_html_deck()
    build_gslides_batch()
