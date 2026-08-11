#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interview Assistant - offline technical interview scoring tool.

Runs fully offline. No internet, no API calls, no login.
Standard library only (tkinter + zipfile) - the Word export is written by hand,
so nothing needs to be pip-installed on the machine that runs this.

Run:      double-click the file (rename to .pyw to hide the console window)
Package:  pyinstaller --onefile --windowed --name InterviewAssistant InterviewAssistant.py
"""

import csv
import hashlib
import os
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox

APP_NAME = "Interview Assistant"
VERSION = "1.0"

# --------------------------------------------------------------------------
# Look and feel
# --------------------------------------------------------------------------

INK = "#14181F"
MUTED = "#6A7480"
FAINT = "#98A2AE"
CANVAS = "#EDF0F2"
SURFACE = "#FFFFFF"
RULE = "#DCE2E6"
HEADER = "#12202B"
HEADER_SOFT = "#22313E"
ACCENT = "#00786C"
ACCENT_DK = "#005F55"
ACCENT_LT = "#7FD4C7"
WARN = "#B4643C"

# single-hue strength ramp, 1..5 - reads as intensity, not traffic lights
SCORE_FILL = {
    1: "#CFE2DF",
    2: "#9CCCC5",
    3: "#63B2A8",
    4: "#2D9587",
    5: "#00786C",
}
SCORE_TEXT = {1: INK, 2: INK, 3: "#FFFFFF", 4: "#FFFFFF", 5: "#FFFFFF"}
SKIP_FILL = "#C8CFD6"
EMPTY_FILL = "#F2F5F7"


def font(size=10, weight="normal", family="body"):
    fam = {
        "display": "Segoe UI Semibold",
        "body": "Segoe UI",
        "mono": "Consolas",
    }[family]
    if sys.platform == "darwin":
        fam = {"Segoe UI Semibold": "Helvetica Neue", "Segoe UI": "Helvetica Neue",
               "Consolas": "Menlo"}[fam]
    elif not sys.platform.startswith("win"):
        fam = {"Segoe UI Semibold": "DejaVu Sans", "Segoe UI": "DejaVu Sans",
               "Consolas": "DejaVu Sans Mono"}[fam]
    return (fam, int(round(size)), weight)


# --------------------------------------------------------------------------
# Question bank - 15 per topic, ordered easy -> hard
# (level, id, question, expected answer)
# --------------------------------------------------------------------------

def _bank(topic, rows):
    out = []
    for i, (q, a) in enumerate(rows, start=1):
        level = "Basic" if i <= 5 else ("Medium" if i <= 10 else "Advanced")
        out.append({
            "qid": "%s-%02d" % (topic[:3].upper(), i),
            "no": i,
            "level": level,
            "question": q,
            "expected": a,
        })
    return out


QUESTIONS = {
    "SQL": _bank("SQL", [
        ("What does SELECT DISTINCT do, and when would you use it?",
         "Returns only unique rows for the selected columns. Used to remove duplicates, e.g. listing every region that appears in a sales table."),
        ("What is the difference between WHERE and HAVING?",
         "WHERE filters individual rows before grouping. HAVING filters groups after GROUP BY, so it can use aggregates like COUNT(*) > 5."),
        ("Explain primary key vs foreign key.",
         "A primary key uniquely identifies each row and cannot be NULL. A foreign key points at a primary key in another table and enforces referential integrity."),
        ("What is the difference between an INNER JOIN and a LEFT JOIN?",
         "INNER JOIN returns only matching rows from both tables. LEFT JOIN returns all rows from the left table plus matches from the right, with NULLs where there is no match."),
        ("What does GROUP BY do? Give a short example.",
         "Collapses rows into groups so aggregates can be applied per group, e.g. SELECT region, COUNT(*) FROM sales GROUP BY region."),
        ("How do NULLs behave in aggregates? Difference between COUNT(*) and COUNT(column)?",
         "Aggregates ignore NULLs. COUNT(*) counts all rows; COUNT(column) counts only rows where that column is not NULL. Also, NULL = NULL is never true - use IS NULL."),
        ("UNION vs UNION ALL - what is the difference and which is faster?",
         "UNION removes duplicates and therefore sorts/hashes the result. UNION ALL keeps everything and is faster. Both need matching column counts and compatible types."),
        ("What is a subquery? Explain correlated vs non-correlated.",
         "A query nested inside another. Non-correlated runs once independently. Correlated references the outer query and runs per outer row, so it is usually slower."),
        ("Explain window functions. Difference between ROW_NUMBER, RANK and DENSE_RANK.",
         "They compute across a set of rows without collapsing them. ROW_NUMBER always gives unique sequential numbers; RANK leaves gaps after ties; DENSE_RANK does not leave gaps."),
        ("Write logic to find the second highest salary in an Employee table.",
         "Common answers: DENSE_RANK() OVER (ORDER BY salary DESC) filtered to 2, or SELECT MAX(salary) WHERE salary < (SELECT MAX(salary)...), or OFFSET 1 ROWS FETCH NEXT 1. Should mention handling ties and no-second-value cases."),
        ("What is an index? Clustered vs non-clustered, and when can an index hurt?",
         "Structure that speeds up lookups. Clustered defines the physical row order (one per table); non-clustered is a separate structure with pointers. Indexes slow down INSERT/UPDATE/DELETE and consume storage."),
        ("What is a CTE? When would you use a recursive one?",
         "A named temporary result set defined with WITH, used to make complex queries readable. Recursive CTEs walk hierarchies such as employee-manager chains or bills of material."),
        ("How would you find duplicate rows and delete all but one?",
         "GROUP BY the key columns HAVING COUNT(*) > 1 to find them; delete using ROW_NUMBER() OVER (PARTITION BY key ORDER BY something) and removing rows where rn > 1."),
        ("A query that used to run in seconds now takes minutes. Walk me through your approach.",
         "Check the execution plan, look for table scans vs seeks, stale statistics, missing or unusable indexes, non-SARGable predicates (functions on columns), implicit conversions, parameter sniffing, data volume growth, and blocking."),
        ("Explain transactions and ACID. What are isolation levels and deadlocks?",
         "ACID = Atomicity, Consistency, Isolation, Durability. Isolation levels (Read Uncommitted through Serializable, plus snapshot) trade concurrency against dirty/non-repeatable/phantom reads. A deadlock is two sessions each holding a lock the other needs; the engine kills one as victim."),
    ]),
    "Python": _bank("Python", [
        ("What is the difference between a list and a tuple?",
         "Lists are mutable and use []; tuples are immutable and use (). Tuples can be dictionary keys and are slightly faster/lighter."),
        ("What are dictionaries and sets, and when would you use each?",
         "Dict stores key-value pairs with fast lookup by key. Set stores unique unordered items, good for membership tests and de-duplication. Both are roughly O(1) lookup."),
        ("How do you loop over a list with its index?",
         "Use enumerate(items) - for i, item in enumerate(items). Should not be manually indexing with range(len(items)) as the first answer."),
        ("What are *args and **kwargs?",
         "*args collects extra positional arguments as a tuple; **kwargs collects extra keyword arguments as a dict. Used for flexible function signatures and pass-through wrappers."),
        ("How do you read a file safely in Python?",
         "with open(path, 'r', encoding='utf-8') as f: - the with block closes the file even on error. Should mention encoding and reading line by line for large files."),
        ("What is a list comprehension? Rewrite a simple loop as one.",
         "[x*2 for x in nums if x > 0] - concise transform + filter in one expression. Faster and more readable than append loops for simple cases."),
        ("Explain mutable vs immutable, and the mutable default argument trap.",
         "Lists/dicts/sets are mutable; str/int/tuple are not. def f(x, items=[]) reuses the same list across calls - use items=None and create inside the function."),
        ("How does exception handling work? What do else and finally do?",
         "try/except catches errors; else runs only if no exception; finally always runs for cleanup. Should mention catching specific exceptions rather than bare except."),
        ("Shallow copy vs deep copy?",
         "Shallow copy (copy.copy, list slicing) copies the outer container but shares nested objects. copy.deepcopy recursively copies everything. Matters for nested lists/dicts."),
        ("What is a decorator? Write one that times a function.",
         "A function that wraps another to add behaviour. Uses an inner wrapper, returns it, and applies with @. Good answer uses functools.wraps and *args/**kwargs."),
        ("Generators vs lists - what does yield actually do?",
         "A generator produces values lazily one at a time instead of building the whole list in memory. yield suspends and resumes the function. Essential for large files/streams."),
        ("Explain classes: __init__, self, inheritance, and one dunder method.",
         "__init__ initialises an instance; self is the instance reference; inheritance reuses a base class and can override methods. Examples of dunders: __str__, __repr__, __len__, __eq__."),
        ("In pandas, when do you use merge vs concat? How does groupby().agg() work?",
         "merge joins on keys like a SQL join; concat stacks frames along an axis. groupby().agg() splits by key, applies one or more aggregations per column, and returns a combined frame."),
        ("What is the GIL? When do you use threading vs multiprocessing?",
         "The Global Interpreter Lock allows only one thread to execute Python bytecode at a time. Threads still help for I/O-bound work; multiprocessing (separate processes) is needed for CPU-bound work."),
        ("A script that processes a large file is too slow. How do you find and fix the bottleneck?",
         "Measure first: cProfile/timeit/line profiler. Then fix - stream instead of loading everything, use vectorised pandas/numpy instead of row loops, avoid repeated work in loops, use sets/dicts for lookups, chunk the file, consider multiprocessing."),
    ]),
    "Power BI": _bank("Power BI", [
        ("What is Power BI, and what is the difference between Desktop and Service?",
         "Desktop is the authoring tool where you model data and build reports. Service is the cloud platform for publishing, sharing, dashboards, and scheduled refresh."),
        ("Which data sources have you connected to, and how?",
         "Get Data connectors - Excel/CSV, SQL Server, SharePoint, web, folder. Should describe a real source they have used and the basic connection steps."),
        ("What is Power Query used for?",
         "Extract and transform before loading: remove columns, filter rows, change types, split/merge, unpivot, append and merge queries. Steps are recorded and repeat on every refresh."),
        ("What is the difference between a report and a dashboard?",
         "A report is multi-page and interactive, built on one dataset. A dashboard is a single Service-only canvas of pinned tiles that can come from several reports."),
        ("How do you choose between a bar chart, a line chart and a card?",
         "Bar/column for comparing categories, line for trends over time, card for a single KPI. Should mention avoiding pie charts with many categories."),
        ("What is a data model? Explain relationships and cardinality.",
         "Tables linked by keys. Cardinality is one-to-many (normal), many-to-many (avoid where possible), one-to-one. Also filter direction - single vs bidirectional."),
        ("Star schema vs one flat table - why does it matter?",
         "Star schema (fact table plus dimension tables) compresses better, gives cleaner relationships, faster DAX and reusable filters. Flat tables duplicate data and break time intelligence."),
        ("Calculated column vs measure - what is the real difference?",
         "A calculated column is computed row by row at refresh and stored in the model (uses memory). A measure is computed at query time based on the current filter context. Use measures for aggregations."),
        ("Explain SUM vs SUMX, and what CALCULATE does.",
         "SUM aggregates one column. SUMX iterates a table row by row evaluating an expression, e.g. SUMX(Sales, Sales[Qty]*Sales[Price]). CALCULATE evaluates an expression with modified filter context - the core DAX function."),
        ("What is the difference between row context and filter context?",
         "Row context is the current row during an iterator or calculated column. Filter context is the set of filters applied by visuals, slicers and CALCULATE. Context transition via CALCULATE turns row context into filter context."),
        ("How do you build a YTD or same-period-last-year measure?",
         "Needs a marked date table with continuous dates. TOTALYTD or CALCULATE(SUM(...), DATESYTD(Date[Date])), and SAMEPERIODLASTYEAR or DATEADD for prior year comparisons."),
        ("Import vs DirectQuery vs Live connection - trade-offs?",
         "Import is fastest and supports full DAX but data is a snapshot needing refresh. DirectQuery queries the source live - fresher, but slower and with DAX/modelling limits. Live connection points at an existing dataset or Analysis Services model."),
        ("How would you implement row-level security?",
         "Define roles with DAX filters on dimension tables, e.g. [Region] = USERPRINCIPALNAME() lookup. Test with View As Role, then assign users or security groups to the role in the Service."),
        ("A report is slow to load. How do you diagnose and fix it?",
         "Performance Analyzer to see which visual and whether time is DAX or visual rendering; DAX Studio for query plans. Fixes - reduce visuals per page, remove high-cardinality columns, push transforms upstream, avoid bidirectional filters and complex measures, aggregate tables, disable auto date/time."),
        ("How do you handle deployment across dev/test/prod, and scheduled refresh of on-prem data?",
         "Workspaces plus deployment pipelines with parameterised data sources. On-prem sources need an on-premises data gateway with stored credentials, then a refresh schedule and failure notifications. Version control via .pbip/Git where available."),
    ]),
    "AI": _bank("AI", [
        ("What is the difference between AI, machine learning and deep learning?",
         "AI is the broad field of machines doing tasks that need intelligence. ML is a subset that learns patterns from data instead of hard-coded rules. Deep learning is a subset of ML using multi-layer neural networks."),
        ("Supervised vs unsupervised learning - give one example of each.",
         "Supervised uses labelled data, e.g. predicting churn from past labelled customers. Unsupervised finds structure without labels, e.g. customer segmentation with clustering."),
        ("What do the terms features, labels and training data mean?",
         "Features are the input variables, labels are the known answers being predicted, and training data is the labelled set the model learns from."),
        ("Classification vs regression?",
         "Classification predicts a discrete category (spam / not spam). Regression predicts a continuous number (next month's revenue)."),
        ("Explain what a large language model is in plain terms.",
         "A model trained on huge amounts of text to predict the next token, which lets it generate and transform language. It is a statistical pattern learner, not a knowledge database."),
        ("What is overfitting? How do you spot and reduce it?",
         "The model memorises the training data and fails on new data - high training accuracy, low validation accuracy. Fix with more data, simpler models, regularisation, dropout, early stopping, cross-validation."),
        ("Why split data into train/validation/test? What is cross-validation?",
         "Train fits the model, validation tunes hyperparameters, test gives an unbiased final estimate. K-fold cross-validation rotates the validation split to use limited data efficiently."),
        ("When is accuracy a misleading metric? What would you use instead?",
         "On imbalanced data - 99% accuracy is trivial if 99% of cases are negative. Use precision, recall, F1, ROC-AUC or PR-AUC, and a confusion matrix; choose based on the cost of false positives vs false negatives."),
        ("What is prompt engineering? Zero-shot vs few-shot?",
         "Structuring the instruction to get reliable output - role, context, constraints, output format. Zero-shot gives no examples; few-shot includes worked examples to demonstrate the pattern."),
        ("What are embeddings and how is similarity measured?",
         "Numeric vector representations where similar meanings sit close together. Similarity is usually cosine distance. Used for search, clustering, deduplication and recommendations."),
        ("Explain RAG and why organisations use it.",
         "Retrieval-Augmented Generation: retrieve relevant documents from a vector or keyword store, then pass them to the model as context. Grounds answers in current, private data without retraining, and enables citations."),
        ("What causes hallucination and how do you reduce it?",
         "The model generates plausible text without grounding. Reduce with retrieval grounding, asking for citations, constrained/structured output, lower temperature, saying 'I don't know' as an allowed answer, validation layers and human review."),
        ("Fine-tuning vs RAG vs prompting - how do you choose?",
         "Prompting first - cheapest. RAG when the issue is missing or changing knowledge. Fine-tuning when the issue is consistent style, format or a narrow task, and you have good training examples. They combine."),
        ("What risks would you raise before deploying an AI feature on company data?",
         "Data privacy and where data is processed, PII handling, bias in training data and outcomes, explainability, over-reliance by users, model drift, audit trail, and a human-in-the-loop for consequential decisions."),
        ("Design an AI solution for a business problem end to end - how would you measure success?",
         "Define the decision it supports and the baseline; data sourcing and quality; approach selection; offline evaluation with a held-out set and a metric tied to business cost; pilot with human review; deployment, monitoring for drift, feedback loop, and a measurable outcome such as hours saved or error rate reduced."),
    ]),
    "Excel": _bank("Excel", [
        ("What is the difference between a relative and an absolute reference?",
         "A1 shifts when copied; $A$1 stays fixed. Mixed refs like A$1 lock only row or column. F4 toggles them."),
        ("Explain VLOOKUP and its main limitations.",
         "VLOOKUP(lookup, table, col_index, FALSE) searches the first column only. Limits - cannot look left, breaks when columns are inserted, slow on large ranges, needs exact-match FALSE."),
        ("How do you write an IF, and how do you handle several conditions?",
         "IF(test, value_if_true, value_if_false). Multiple conditions via nested IF, IFS, or IF with AND/OR."),
        ("What do SUMIF, SUMIFS and COUNTIFS do?",
         "SUMIF sums on one condition; SUMIFS on multiple; COUNTIFS counts on multiple. Criteria ranges must be the same size, and wildcards/operators like \">100\" are allowed."),
        ("How would you remove duplicates and split a full name into two columns?",
         "Data > Remove Duplicates for de-duplication. Text to Columns with a space delimiter, or Flash Fill, or LEFT/RIGHT with FIND for the split."),
        ("Walk me through building a pivot table to summarise sales by region and month.",
         "Format source as a table, Insert > PivotTable, Region to Rows, dates to Columns grouped by month, Sales to Values as Sum. Should mention refresh after data changes."),
        ("INDEX/MATCH or XLOOKUP - why prefer them over VLOOKUP?",
         "INDEX(return_range, MATCH(lookup, lookup_range, 0)) can look in any direction and survives column insertion. XLOOKUP does the same in one function with a built-in if-not-found argument."),
        ("How do you apply conditional formatting driven by a formula?",
         "New Rule > Use a formula. The formula is written for the top-left cell of the range with the right relative/absolute mix, e.g. =$C2<TODAY(), and applies across the range."),
        ("How do you set up a dropdown and stop bad data entry?",
         "Data Validation > List, pointing at a named range or table column. Add input and error messages; combine with validation rules for dates/numbers."),
        ("How do you handle errors like #N/A and #REF?",
         "Wrap in IFERROR or IFNA for controlled output. #N/A means lookup not found, #REF means a referenced cell was deleted, #VALUE means wrong data type. Should mention not blanket-hiding real errors."),
        ("What is Power Query in Excel and when do you use it?",
         "Repeatable ETL inside Excel - import, clean, unpivot, append and merge queries, then refresh with one click. Replaces manual copy-paste cleanup on recurring files."),
        ("Explain dynamic array functions like FILTER, UNIQUE and SORT.",
         "They spill results across a range from one formula. FILTER returns rows meeting a condition, UNIQUE removes duplicates, SORT orders results. Combined, they replace many helper columns."),
        ("Why use named ranges and Excel tables with structured references?",
         "Tables auto-expand, so formulas and pivots pick up new rows without editing ranges. Structured references like Sales[Amount] are self-documenting and reduce broken ranges."),
        ("When would you use a macro or VBA, and when would you avoid it?",
         "Use for repetitive multi-step tasks that formulas and Power Query cannot do, e.g. formatting and distributing many files. Avoid when Power Query or formulas will do, when the file must be shared broadly, or where macros are blocked by policy."),
        ("You inherit a slow, fragile workbook. How do you stabilise it?",
         "Find volatile functions (INDIRECT, OFFSET, NOW, entire-column refs), broken external links, and array formulas over whole columns. Replace with tables and structured refs, move transforms to Power Query, split calculation from presentation, document assumptions, add validation, and set manual calc while working."),
    ]),
}

TOPICS = list(QUESTIONS.keys())
DECISIONS = ["Strong hire", "Hire", "Hold", "No hire"]


# --------------------------------------------------------------------------
# Minimal .docx writer (no external packages)
# --------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:color w:val="14181F"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:spacing w:after="40"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:b/><w:sz w:val="46"/><w:color w:val="12202B"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:spacing w:before="320" w:after="140"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="27"/><w:color w:val="00786C"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="23"/><w:color w:val="12202B"/></w:rPr></w:style>
</w:styles>""" % W_NS

_DOCUMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="%s"><w:body>%%s
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1100" w:right="1000" w:bottom="1100" w:left="1000" w:header="600" w:footer="600" w:gutter="0"/></w:sectPr>
</w:body></w:document>""" % W_NS


def xesc(text):
    if text is None:
        text = ""
    text = str(text)
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    # strip control characters Word will reject
    return "".join(ch for ch in text if ch in "\n\t" or ord(ch) >= 32)


def _runs(text, bold=False, italic=False, color=None, size=None, mono=False):
    rpr = "<w:rPr>"
    if mono:
        rpr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>'
    if bold:
        rpr += "<w:b/>"
    if italic:
        rpr += "<w:i/>"
    if color:
        rpr += '<w:color w:val="%s"/>' % color
    if size:
        rpr += '<w:sz w:val="%d"/>' % int(size * 2)
    rpr += "</w:rPr>"
    parts = []
    for i, line in enumerate(xesc(text).split("\n")):
        if i:
            parts.append("<w:r>%s<w:br/></w:r>" % rpr)
        parts.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, line))
    return "".join(parts)


class Docx(object):
    """Builds a simple, clean Word document without any third-party package."""

    def __init__(self):
        self.blocks = []

    def para(self, text="", style=None, bold=False, italic=False, color=None,
             size=None, mono=False, space_after=None, shade=None, align=None):
        ppr = ""
        inner = ""
        if style:
            inner += '<w:pStyle w:val="%s"/>' % style
        if align:
            inner += '<w:jc w:val="%s"/>' % align
        if shade:
            inner += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade
            inner += ('<w:pBdr><w:left w:val="single" w:sz="18" w:space="6" w:color="00786C"/></w:pBdr>')
            inner += '<w:ind w:left="140"/>'
        if space_after is not None:
            inner += '<w:spacing w:after="%d"/>' % space_after
        if inner:
            ppr = "<w:pPr>%s</w:pPr>" % inner
        self.blocks.append("<w:p>%s%s</w:p>" % (
            ppr, _runs(text, bold, italic, color, size, mono) if text != "" else ""))
        return self

    def title(self, text):
        return self.para(text, style="Title")

    def h1(self, text):
        return self.para(text, style="Heading1")

    def h2(self, text):
        return self.para(text, style="Heading2")

    def small(self, text, color=MUTED.lstrip("#")):
        return self.para(text, size=9, color=color)

    def bullet(self, text):
        return self.para(u"\u2022  " + text, space_after=60)

    def rule(self):
        self.blocks.append(
            '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" '
            'w:color="DCE2E6"/></w:pBdr><w:spacing w:after="160"/></w:pPr></w:p>')
        return self

    def page_break(self):
        self.blocks.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        return self

    def table(self, header, rows, widths=None, zebra=True):
        ncol = len(header) if header else (len(rows[0]) if rows else 0)
        if not ncol:
            return self
        widths = widths or [int(9600 / ncol)] * ncol
        borders = ("<w:tblBorders>"
                   + "".join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="DCE2E6"/>' % s
                             for s in ["top", "left", "bottom", "right", "insideH", "insideV"])
                   + "</w:tblBorders>")
        xml = ['<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>%s'
               '<w:tblCellMar><w:top w:w="70" w:type="dxa"/><w:left w:w="110" w:type="dxa"/>'
               '<w:bottom w:w="70" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tblCellMar>'
               '</w:tblPr>' % borders]
        xml.append("<w:tblGrid>%s</w:tblGrid>"
                   % "".join('<w:gridCol w:w="%d"/>' % w for w in widths))

        def cell(text, w, fill=None, bold=False, color=None, mono=False, size=None):
            shd = '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill if fill else ""
            return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s'
                    '<w:vAlign w:val="center"/></w:tcPr>'
                    '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>%s</w:p></w:tc>'
                    % (w, shd, _runs(text, bold=bold, color=color, mono=mono, size=size)))

        if header:
            xml.append("<w:tr><w:trPr><w:tblHeader/></w:trPr>")
            for h, w in zip(header, widths):
                xml.append(cell(h, w, fill="12202B", bold=True, color="FFFFFF", size=9))
            xml.append("</w:tr>")
        for i, row in enumerate(rows):
            fill = "F4F7F8" if (zebra and i % 2) else None
            xml.append("<w:tr>")
            for c, w in zip(row, widths):
                xml.append(cell(c, w, fill=fill, size=9.5))
            xml.append("</w:tr>")
        xml.append("</w:tbl>")
        self.blocks.append("".join(xml))
        self.para(space_after=120)
        return self

    def save(self, path):
        body = "".join(self.blocks)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CONTENT_TYPES)
            z.writestr("_rels/.rels", _RELS)
            z.writestr("word/_rels/document.xml.rels", _DOC_RELS)
            z.writestr("word/styles.xml", _STYLES)
            z.writestr("word/document.xml", _DOCUMENT % body)
        return path


def read_text_file(path):
    """Read .txt/.md/.csv/.vtt directly, or pull the text out of a .docx."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
        xml = re.sub(r"</w:p>", "\n", xml)
        xml = re.sub(r"<w:br[^>]*/>", "\n", xml)
        text = re.sub(r"<[^>]+>", "", xml)
        text = (text.replace("&amp;", "&").replace("&lt;", "<")
                    .replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'"))
        return "\n".join(l.rstrip() for l in text.splitlines())
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, LookupError):
            continue
    with open(path, "rb") as f:
        return f.read().decode("utf-8", "ignore")


def anonymize(text, names):
    """Replace real names with role labels. names = {'Real Name': 'Candidate'}"""
    hits = 0
    for real, label in names.items():
        real = (real or "").strip()
        if len(real) < 3:
            continue
        parts = [p for p in re.split(r"\s+", real) if len(p) >= 3]
        for token in sorted(set([real] + parts), key=len, reverse=True):
            pattern = re.compile(r"\b%s\b" % re.escape(token), re.IGNORECASE)
            text, n = pattern.subn(label, text)
            hits += n
    return text, hits


# --------------------------------------------------------------------------
# Loading a question bank from CSV or Excel
# --------------------------------------------------------------------------

COLUMN_ALIASES = {
    "question": ("question", "questions", "questiontext", "q"),
    "expected": ("expectedanswer", "expectedanswers", "expected", "answer",
                 "answerkey", "answers", "description", "briefdescription",
                 "expectedanswerkey"),
    "level": ("level", "difficulty", "complexity", "band"),
    "topic": ("topic", "technology", "tech", "subject", "category", "skill", "area"),
    "qid": ("questionid", "qid", "id", "sno", "slno", "srno", "no"),
}

LEVEL_ALIASES = {
    "Basic": ("basic", "easy", "beginner", "simple", "low", "l1", "1", "b"),
    "Medium": ("medium", "moderate", "intermediate", "average", "mid", "l2", "2", "m"),
    "Advanced": ("advanced", "hard", "difficult", "expert", "complex", "high", "l3", "3", "a"),
}

LEVEL_ORDER = {"Basic": 0, "Medium": 1, "Advanced": 2}

QUESTION_TEMPLATE = (
    "Topic,Level,QuestionID,Question,ExpectedAnswer\r\n"
    "SQL,Basic,SQL-01,What does SELECT DISTINCT do?,"
    "Returns only unique rows for the selected columns.\r\n"
    "SQL,Medium,SQL-02,UNION vs UNION ALL?,"
    "UNION removes duplicates and sorts; UNION ALL keeps everything and is faster.\r\n"
    "SQL,Advanced,SQL-03,How do you tune a slow query?,"
    "Read the execution plan; check indexes; statistics; non-SARGable predicates.\r\n"
)


def norm_key(text):
    return re.sub(r"[^a-z0-9]", "", (text or "").lower())


def col_to_index(ref):
    letters = "".join(ch for ch in (ref or "A") if ch.isalpha()).upper() or "A"
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_xlsx_table(path):
    """Read the first worksheet of an .xlsx into a list of row lists."""
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root:
                shared.append("".join(t.text or "" for t in si.iter(ns + "t")))
        sheets = [n for n in names
                  if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        if not sheets:
            raise ValueError("This Excel file has no worksheet that can be read.")
        sheets.sort(key=lambda n: int(re.sub(r"\D", "", n.rsplit("/", 1)[-1]) or 0))
        root = ET.fromstring(z.read(sheets[0]))
        rows = []
        for r in root.iter(ns + "row"):
            cells = {}
            for c in r.iter(ns + "c"):
                idx = col_to_index(c.get("r"))
                ctype = c.get("t")
                if ctype == "inlineStr":
                    node = c.find(ns + "is")
                    val = "".join(x.text or "" for x in node.iter(ns + "t")) if node is not None else ""
                else:
                    v = c.find(ns + "v")
                    val = (v.text or "") if v is not None else ""
                    if ctype == "s" and val != "":
                        try:
                            val = shared[int(val)]
                        except (ValueError, IndexError):
                            val = ""
                cells[idx] = (val or "").strip()
            if any(v for v in cells.values()):
                rows.append([cells.get(i, "") for i in range(max(cells) + 1)])
        return rows


def read_csv_table(path):
    text = read_text_file(path)
    try:
        dialect = csv.Sniffer().sniff(text[:4000], delimiters=",;\t|")
    except Exception:
        dialect = csv.excel
    rows = [[(c or "").strip() for c in row]
            for row in csv.reader(text.splitlines(), dialect)]
    return [r for r in rows if any(r)]


def normalise_level(value):
    key = norm_key(value)
    for level, aliases in LEVEL_ALIASES.items():
        if key in aliases:
            return level
    return None


def load_questions_from_file(path, topic):
    """Return (questions, topics_found). Raises ValueError with a readable message."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        rows = read_xlsx_table(path)
    elif ext in (".csv", ".txt"):
        rows = read_csv_table(path)
    elif ext == ".xls":
        raise ValueError("Old .xls files are not supported. Open it in Excel and "
                         "save as .xlsx or .csv, then try again.")
    else:
        raise ValueError("Only CSV and Excel files can be used for the question bank.")

    rows = [r for r in rows if any((c or "").strip() for c in r)]
    if len(rows) < 2:
        raise ValueError("That file has no question rows under the header.")

    header = rows[0]
    colmap = {}
    for i, name in enumerate(header):
        key = norm_key(name)
        for field, aliases in COLUMN_ALIASES.items():
            if key in aliases and field not in colmap:
                colmap[field] = i
    if "question" not in colmap:
        raise ValueError(
            "No Question column found.\n\nThe first row must be a header. "
            "Required column: Question.\nOptional columns: Topic, Level, "
            "QuestionID, ExpectedAnswer.\n\nColumns found: %s"
            % (", ".join(h for h in header if h) or "none"))

    def cell(row, field):
        i = colmap.get(field)
        return (row[i].strip() if i is not None and i < len(row) else "")

    topics_found, raw = [], []
    for row in rows[1:]:
        text = cell(row, "question")
        if not text:
            continue
        rec_topic = cell(row, "topic")
        if rec_topic and rec_topic not in topics_found:
            topics_found.append(rec_topic)
        raw.append({
            "question": text,
            "expected": cell(row, "expected") or "No answer key supplied in the file.",
            "level": normalise_level(cell(row, "level")),
            "topic": rec_topic,
            "qid": cell(row, "qid"),
        })

    if "topic" in colmap and topics_found:
        wanted = norm_key(topic)
        picked = [q for q in raw if norm_key(q["topic"]) == wanted]
        if not picked:
            raise ValueError(
                "No rows in that file have the topic \"%s\".\n\nTopics in the file: %s"
                % (topic, ", ".join(topics_found)))
    else:
        picked = raw

    # fill in any missing levels by splitting the list into three
    if any(q["level"] is None for q in picked):
        n = len(picked)
        for i, q in enumerate(picked):
            if q["level"] is None:
                q["level"] = ("Basic" if i < n / 3.0
                              else ("Medium" if i < 2 * n / 3.0 else "Advanced"))

    picked.sort(key=lambda q: LEVEL_ORDER[q["level"]])
    out = []
    for i, q in enumerate(picked, start=1):
        out.append({
            "qid": q["qid"] or "%s-%02d" % (norm_key(topic)[:3].upper() or "QST", i),
            "no": i,
            "level": q["level"],
            "question": q["question"],
            "expected": q["expected"],
        })
    return out, topics_found


# --------------------------------------------------------------------------
# Widgets
# --------------------------------------------------------------------------

def card(parent, pad=22, bg=SURFACE, border=RULE):
    outer = tk.Frame(parent, bg=border)
    inner = tk.Frame(outer, bg=bg, padx=pad, pady=pad)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    return outer, inner


class Button(tk.Frame):
    KINDS = {
        "primary": (ACCENT, "#FFFFFF", ACCENT_DK, ACCENT),
        "dark":    (HEADER, "#FFFFFF", HEADER_SOFT, HEADER),
        "ghost":   (SURFACE, INK, "#F1F5F6", RULE),
        "quiet":   (CANVAS, MUTED, "#E2E7EA", CANVAS),
    }

    def __init__(self, parent, text, command=None, kind="primary", pady=10, padx=20,
                 size=10, bg=None):
        base, fg, hover, border = self.KINDS[kind]
        tk.Frame.__init__(self, parent, bg=border)
        self.base, self.hover, self.command, self.enabled = base, hover, command, True
        self.lbl = tk.Label(self, text=text, bg=base, fg=fg, padx=padx, pady=pady,
                            font=font(size, "bold", "display" if kind in ("primary", "dark") else "body"),
                            cursor="hand2")
        self.lbl.pack(fill="both", expand=True, padx=1, pady=1)
        for w in (self, self.lbl):
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", self._on_click)

    def _on_enter(self, _=None):
        if self.enabled:
            self.lbl.configure(bg=self.hover)

    def _on_leave(self, _=None):
        if self.enabled:
            self.lbl.configure(bg=self.base)

    def _on_click(self, _=None):
        if self.enabled and self.command:
            self.command()

    def set_text(self, text):
        self.lbl.configure(text=text)

    def set_enabled(self, on):
        self.enabled = on
        self.lbl.configure(bg=self.base if on else "#D9DFE3",
                           fg="#FFFFFF" if on else FAINT,
                           cursor="hand2" if on else "arrow")


class Field(tk.Frame):
    """Labelled single-line input with placeholder text."""

    def __init__(self, parent, label, placeholder="", bg=SURFACE, width=None):
        tk.Frame.__init__(self, parent, bg=bg)
        tk.Label(self, text=label.upper(), bg=bg, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 5))
        box = tk.Frame(self, bg=RULE)
        box.pack(fill="x")
        self.var = tk.StringVar()
        self.entry = tk.Entry(box, textvariable=self.var, bd=0, relief="flat", bg=SURFACE,
                              fg=INK, insertbackground=ACCENT, font=font(10), width=width or 18)
        self.entry.pack(fill="x", ipady=7, ipadx=9, padx=1, pady=1)
        self.placeholder = placeholder
        self._ph_on = False
        if placeholder:
            self._show_ph()
            self.entry.bind("<FocusIn>", self._focus_in)
            self.entry.bind("<FocusOut>", self._focus_out)
        self.entry.bind("<FocusIn>", lambda e: box.configure(bg=ACCENT), add="+")
        self.entry.bind("<FocusOut>", lambda e: box.configure(bg=RULE), add="+")

    def _show_ph(self):
        self.var.set(self.placeholder)
        self.entry.configure(fg=FAINT)
        self._ph_on = True

    def _focus_in(self, _):
        if self._ph_on:
            self.var.set("")
            self.entry.configure(fg=INK)
            self._ph_on = False

    def _focus_out(self, _):
        if not self.var.get().strip():
            self._show_ph()

    def get(self):
        return "" if self._ph_on else self.var.get().strip()

    def set(self, value):
        self._ph_on = False
        self.entry.configure(fg=INK)
        self.var.set(value)


class TextBox(tk.Frame):
    def __init__(self, parent, label, height=5, bg=SURFACE, hint=None):
        tk.Frame.__init__(self, parent, bg=bg)
        row = tk.Frame(self, bg=bg)
        row.pack(fill="x", pady=(0, 5))
        tk.Label(row, text=label.upper(), bg=bg, fg=MUTED, font=font(8, "bold")).pack(side="left")
        if hint:
            tk.Label(row, text=hint, bg=bg, fg=FAINT, font=font(8)).pack(side="right")
        box = tk.Frame(self, bg=RULE)
        box.pack(fill="both", expand=True)
        self.text = tk.Text(box, height=height, bd=0, relief="flat", bg=SURFACE, fg=INK,
                            insertbackground=ACCENT, font=font(10), wrap="word",
                            padx=10, pady=8, spacing1=1, spacing3=3)
        self.text.pack(fill="both", expand=True, padx=1, pady=1)
        self.text.bind("<FocusIn>", lambda e: box.configure(bg=ACCENT))
        self.text.bind("<FocusOut>", lambda e: box.configure(bg=RULE))

    def get(self):
        return self.text.get("1.0", "end").strip()

    def set(self, value):
        self.text.delete("1.0", "end")
        self.text.insert("1.0", value or "")


class ScoreBar(tk.Frame):
    """1-5 selector. Labels sit under the buttons so the scale is unambiguous."""
    ANCHORS = {1: "Poor", 3: "Adequate", 5: "Excellent"}

    def __init__(self, parent, command=None, bg=SURFACE):
        tk.Frame.__init__(self, parent, bg=bg)
        self.command = command
        self.value = None
        self.cells = {}
        row = tk.Frame(self, bg=bg)
        row.pack(anchor="w")
        for n in range(1, 6):
            wrap = tk.Frame(row, bg=RULE)
            wrap.pack(side="left", padx=(0, 8))
            lbl = tk.Label(wrap, text=str(n), bg=SURFACE, fg=INK, width=3,
                           font=font(15, "bold", "display"), cursor="hand2", pady=8)
            lbl.pack(padx=1, pady=1)
            lbl.bind("<Button-1>", lambda e, v=n: self.set(v))
            lbl.bind("<Enter>", lambda e, v=n: self._hover(v, True))
            lbl.bind("<Leave>", lambda e, v=n: self._hover(v, False))
            self.cells[n] = (wrap, lbl)
        tk.Label(self, text="1  poor      3  adequate      5  excellent", bg=bg,
                 fg=FAINT, font=font(8)).pack(anchor="w", pady=(7, 0))

    def _hover(self, n, on):
        if self.value == n:
            return
        self.cells[n][1].configure(bg="#EDF4F3" if on else SURFACE)

    def set(self, value, fire=True):
        self.value = value
        for n, (wrap, lbl) in self.cells.items():
            if n == value:
                wrap.configure(bg=ACCENT)
                lbl.configure(bg=SCORE_FILL[n], fg=SCORE_TEXT[n])
            else:
                wrap.configure(bg=RULE)
                lbl.configure(bg=SURFACE, fg=INK)
        if fire and self.command:
            self.command(value)

    def clear(self):
        self.value = None
        for wrap, lbl in self.cells.values():
            wrap.configure(bg=RULE)
            lbl.configure(bg=SURFACE, fg=INK)


class Scroll(tk.Frame):
    def __init__(self, parent, bg=CANVAS):
        tk.Frame.__init__(self, parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.vsb = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview,
                                width=11, bd=0, relief="flat", highlightthickness=0,
                                troughcolor=CANVAS, bg="#C3CDD4", activebackground=MUTED)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._resize)
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfig(self._win, width=e.width))
        self.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._wheel))
        self.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _resize(self, _=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _wheel(self, event):
        delta = -1 * int(event.delta / 120) if event.delta else 0
        self.canvas.yview_scroll(delta, "units")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

ID_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"   # no I, O, 0, 1


def make_interview_id(name, cid):
    seed = "%s|%s|%s" % (name, cid, datetime.now().isoformat(timespec="microseconds"))
    n = int(hashlib.sha1(seed.encode("utf-8")).hexdigest()[:14], 16)
    out = ""
    for _ in range(6):
        out += ID_ALPHABET[n % len(ID_ALPHABET)]
        n //= len(ID_ALPHABET)
    return "IV-%s-%s" % (out[:3], out[3:])


def safe_name(text):
    text = re.sub(r"[^A-Za-z0-9 _.-]", "", (text or "").strip())
    return re.sub(r"\s+", "_", text)[:40] or "Candidate"


def open_folder(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                      # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def default_output_dir():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    base = desktop if os.path.isdir(desktop) else os.path.expanduser("~")
    return os.path.join(base, "Interview Records")


# --------------------------------------------------------------------------
# Application
# --------------------------------------------------------------------------

STEPS = ["Setup", "Interview", "Wrap-up", "Documents"]


class App(tk.Tk):

    def __init__(self):
        tk.Tk.__init__(self)
        self.title("%s %s" % (APP_NAME, VERSION))
        self.configure(bg=CANVAS)
        self.minsize(1040, 680)
        self._centre(1200, 800)
        self.step = 0
        self.reset_data()
        self._build_chrome()
        self.show_setup()

    # -- window ------------------------------------------------------------
    def _centre(self, w, h):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(w, sw - 80), min(h, sh - 100)
        self.geometry("%dx%d+%d+%d" % (w, h, (sw - w) // 2, max(0, (sh - h) // 2 - 20)))

    def reset_data(self):
        self.data = {
            "interview_id": None,
            "candidate": {},
            "topic": None,
            "questions": [],
            "responses": {},
            "rating": None,
            "comment": "",
            "decision": None,
            "started": None,
            "finished": None,
            "transcript_source": None,
            "anonymised": False,
            "source_note": "",
        }
        self.q_index = 0

    # -- chrome ------------------------------------------------------------
    def _build_chrome(self):
        bar = tk.Frame(self, bg=HEADER, height=64)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        left = tk.Frame(bar, bg=HEADER)
        left.pack(side="left", padx=26)
        tk.Label(left, text=APP_NAME, bg=HEADER, fg="#FFFFFF",
                 font=font(14, "bold", "display")).pack(side="left", pady=16)
        self.id_chip = tk.Label(left, text="", bg=HEADER, fg=ACCENT_LT, font=font(10, "bold", "mono"))
        self.id_chip.pack(side="left", padx=(16, 0), pady=18)

        self.step_bar = tk.Frame(bar, bg=HEADER)
        self.step_bar.pack(side="right", padx=26)
        self.step_labels = []
        for i, name in enumerate(STEPS):
            if i:
                tk.Label(self.step_bar, text="\u2022", bg=HEADER, fg="#3C4B58",
                         font=font(9)).pack(side="left", padx=8, pady=22)
            lbl = tk.Label(self.step_bar, text="%d  %s" % (i + 1, name), bg=HEADER,
                           fg="#6C7C89", font=font(9, "bold"))
            lbl.pack(side="left", pady=22)
            self.step_labels.append(lbl)

        tk.Frame(self, bg=ACCENT, height=2).pack(fill="x")
        self.body = tk.Frame(self, bg=CANVAS)
        self.body.pack(fill="both", expand=True)

    def _set_step(self, i):
        self.step = i
        for n, lbl in enumerate(self.step_labels):
            lbl.configure(fg=ACCENT_LT if n == i else ("#93A2AE" if n < i else "#5B6B78"))

    def _clear(self):
        for w in self.body.winfo_children():
            w.destroy()

    def _set_id_chip(self):
        iid = self.data["interview_id"]
        self.id_chip.configure(text=("INTERVIEW  " + iid) if iid else "")

    # ==================================================================
    # 1. SETUP
    # ==================================================================
    def show_setup(self):
        self._clear()
        self._set_step(0)
        scroll = Scroll(self.body)
        scroll.pack(fill="both", expand=True)
        page = tk.Frame(scroll.inner, bg=CANVAS, padx=34, pady=26)
        page.pack(fill="both", expand=True)

        tk.Label(page, text="Set up the interview", bg=CANVAS, fg=INK,
                 font=font(20, "bold", "display")).pack(anchor="w")
        tk.Label(page, text="Enter the candidate's details, pick a topic, and load or upload the questions.",
                 bg=CANVAS, fg=MUTED, font=font(10)).pack(anchor="w", pady=(4, 20))

        # --- candidate details ---
        outer, c = card(page)
        outer.pack(fill="x", pady=(0, 18))
        tk.Label(c, text="Candidate details", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w", pady=(0, 14))

        grid = tk.Frame(c, bg=SURFACE)
        grid.pack(fill="x")
        for col in range(4):
            grid.columnconfigure(col, weight=1, uniform="f")

        now = datetime.now()
        specs = [
            ("team", "Team name", "e.g. Data Services"),
            ("candidate", "Candidate name", "Full name"),
            ("cid", "Candidate ID", "e.g. CAND001"),
            ("req", "Requisition ID", "e.g. REQ-4821"),
            ("grade", "Grade", "e.g. 110 / 120"),
            ("role", "Role", "e.g. Data Analyst"),
            ("date", "Interview date", ""),
            ("time", "Interview time", ""),
            ("interviewer", "Interviewer name", "Your name"),
        ]
        self.fields = {}
        for i, (key, label, ph) in enumerate(specs):
            f = Field(grid, label, ph)
            f.grid(row=i // 4, column=i % 4, sticky="ew", padx=(0, 14), pady=(0, 14))
            self.fields[key] = f
        self.fields["date"].set(now.strftime("%d %b %Y"))
        self.fields["time"].set(now.strftime("%I:%M %p").lstrip("0"))
        prev = self.data.get("candidate") or {}
        for k, v in prev.items():
            if k in self.fields and v:
                self.fields[k].set(v)

        tk.Label(c, text="Candidate name and interviewer name are required. Everything else is optional.",
                 bg=SURFACE, fg=FAINT, font=font(9)).pack(anchor="w")

        # --- topic ---
        outer2, t = card(page)
        outer2.pack(fill="x", pady=(0, 18))
        tk.Label(t, text="Topic", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w")
        tk.Label(t, text="Pick the technology, then load the built-in set of 15 questions "
                         "or upload your own bank as CSV or Excel.",
                 bg=SURFACE, fg=MUTED, font=font(9.5)).pack(anchor="w", pady=(3, 14))

        tiles = tk.Frame(t, bg=SURFACE)
        tiles.pack(anchor="w")
        self.topic_var = tk.StringVar(value=self.data.get("topic") or "")
        self.topic_tiles = {}
        for topic in TOPICS:
            wrap = tk.Frame(tiles, bg=RULE)
            wrap.pack(side="left", padx=(0, 10))
            lbl = tk.Label(wrap, text=topic, bg=SURFACE, fg=INK, font=font(11, "bold"),
                           padx=22, pady=13, cursor="hand2", width=9)
            lbl.pack(padx=1, pady=1)
            lbl.bind("<Button-1>", lambda e, tp=topic: self._pick_topic(tp))
            self.topic_tiles[topic] = (wrap, lbl)

        srow = tk.Frame(t, bg=SURFACE)
        srow.pack(fill="x", pady=(18, 0))
        self.btn_load = Button(srow, "Load built-in questions", self._load_questions,
                               kind="ghost")
        self.btn_load.pack(side="left")
        Button(srow, "Upload questions  \u00b7  CSV or Excel", self._upload_questions,
               kind="ghost").pack(side="left", padx=10)
        Button(srow, "Save a blank template", self._save_template,
               kind="ghost", size=9, pady=11).pack(side="left")

        self.source_lbl = tk.Label(t, text="No questions loaded yet.", bg=SURFACE,
                                   fg=FAINT, font=font(9.5), anchor="w", justify="left",
                                   wraplength=820)
        self.source_lbl.pack(anchor="w", pady=(12, 0))

        self.counts = tk.Frame(t, bg=SURFACE)
        self.counts.pack(fill="x", pady=(14, 0))
        self.count_labels = {}
        for name in ("Total", "Basic", "Medium", "Advanced"):
            box_o, box = card(self.counts, pad=12, bg="#F5F8F9", border="#E6ECEE")
            box_o.pack(side="left", padx=(0, 10))
            tk.Label(box, text=name.upper(), bg="#F5F8F9", fg=MUTED,
                     font=font(8, "bold")).pack()
            v = tk.Label(box, text="0", bg="#F5F8F9", fg=ACCENT, font=font(19, "bold", "display"),
                         width=6)
            v.pack()
            self.count_labels[name] = v

        row = tk.Frame(page, bg=CANVAS)
        row.pack(fill="x")
        self.btn_start = Button(row, "Start interview  \u2192", self._start_interview)
        self.btn_start.pack(side="left")
        self.btn_start.set_enabled(False)

        if self.data.get("topic"):
            self._pick_topic(self.data["topic"])
        if self.data.get("questions"):
            self._refresh_counts(self.data["questions"])
            self.source_lbl.configure(text=self.data.get("source_note", ""), fg=INK)
            self.btn_start.set_enabled(True)

    def _pick_topic(self, topic):
        self.topic_var.set(topic)
        for name, (wrap, lbl) in self.topic_tiles.items():
            on = name == topic
            wrap.configure(bg=ACCENT if on else RULE)
            lbl.configure(bg=ACCENT if on else SURFACE, fg="#FFFFFF" if on else INK)

    def _refresh_counts(self, qs):
        self.count_labels["Total"].configure(text=str(len(qs)))
        for lvl in ("Basic", "Medium", "Advanced"):
            self.count_labels[lvl].configure(
                text=str(len([q for q in qs if q["level"] == lvl])))

    def _collect_candidate(self):
        return {k: f.get() for k, f in self.fields.items()}

    def _ready_to_load(self):
        """Topic chosen and the two required names filled in."""
        if not self.topic_var.get():
            messagebox.showinfo("Pick a topic", "Choose a topic first.")
            return False
        cand = self._collect_candidate()
        if not cand["candidate"] or not cand["interviewer"]:
            messagebox.showinfo("Details needed",
                                "Enter the candidate name and the interviewer name first.")
            return False
        self.data["candidate"] = cand
        self.data["topic"] = self.topic_var.get()
        if not self.data["interview_id"]:
            self.data["interview_id"] = make_interview_id(cand["candidate"], cand["cid"])
        self._set_id_chip()
        return True

    def _apply_questions(self, questions, note):
        self.data["questions"] = questions
        self.data["responses"] = {}
        self.data["source_note"] = note
        self._refresh_counts(questions)
        self.source_lbl.configure(text=note, fg=INK)
        self.btn_start.set_enabled(True)

    def _load_questions(self):
        if not self._ready_to_load():
            return
        topic = self.data["topic"]
        self._apply_questions([dict(q) for q in QUESTIONS[topic]],
                              "Loaded the built-in set of 15 %s questions." % topic)
        self.btn_load.set_text("Reload built-in questions")

    def _upload_questions(self):
        if not self._ready_to_load():
            return
        topic = self.data["topic"]
        path = filedialog.askopenfilename(
            title="Choose the question bank for %s" % topic,
            filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"),
                       ("Excel workbook", "*.xlsx *.xlsm"),
                       ("CSV file", "*.csv")])
        if not path:
            return
        if os.path.splitext(path)[1].lower() not in (".xlsx", ".xlsm", ".csv"):
            messagebox.showinfo("Wrong file type",
                                "Only CSV and Excel files can be used. Save your "
                                "question bank as .xlsx or .csv and try again.")
            return
        try:
            questions, topics = load_questions_from_file(path, topic)
        except ValueError as exc:
            messagebox.showwarning("Could not read that file", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Could not read that file",
                                 "%s\n\nCheck the file opens in Excel and try again."
                                 % exc)
            return
        if not questions:
            messagebox.showwarning("No questions found",
                                   "No usable question rows were found in that file.")
            return
        note = "Loaded %d %s question%s from %s." % (
            len(questions), topic, "" if len(questions) == 1 else "s",
            os.path.basename(path))
        if topics and len(topics) > 1:
            note += "  Other topics in the file: %s." % ", ".join(
                t for t in topics if norm_key(t) != norm_key(topic))
        self._apply_questions(questions, note)
        self.btn_load.set_text("Use built-in questions instead")

    def _save_template(self):
        path = filedialog.asksaveasfilename(
            title="Save the question template",
            defaultextension=".csv", initialfile="Question_Bank_Template.csv",
            filetypes=[("CSV file", "*.csv")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write(QUESTION_TEMPLATE)
        except Exception as exc:
            messagebox.showerror("Could not save the template", str(exc))
            return
        messagebox.showinfo(
            "Template saved",
            "Fill in the rows and upload the file, or save it as .xlsx first.\n\n"
            "Required column: Question.\nOptional: Topic, Level, QuestionID, "
            "ExpectedAnswer.\n\nLevel accepts Basic, Medium or Advanced. Rows are "
            "sorted easy to hard, and only rows matching the topic you picked are used.")

    def _start_interview(self):
        self.data["candidate"] = self._collect_candidate()
        self.data["started"] = datetime.now().isoformat(timespec="seconds")
        self.q_index = 0
        self.show_interview()

    # ==================================================================
    # 2. INTERVIEW
    # ==================================================================
    def _response(self, idx):
        return self.data["responses"].setdefault(
            idx, {"score": None, "skipped": False, "notes": ""})

    def _attempted(self):
        out = []
        for idx, r in self.data["responses"].items():
            if r["skipped"]:
                out.append((idx, 0))
            elif r["score"] is not None:
                out.append((idx, r["score"]))
        return out

    def stats(self):
        att = self._attempted()
        total = len(self.data["questions"]) or 1
        scored = [s for _, s in att]
        avg = (sum(scored) / len(scored)) if scored else 0.0
        by_level = {}
        for lvl in ("Basic", "Medium", "Advanced"):
            vals = [s for i, s in att if self.data["questions"][i]["level"] == lvl]
            by_level[lvl] = (sum(vals) / len(vals)) if vals else None
        return {
            "attempted": len(att),
            "total": total,
            "skipped": len([1 for r in self.data["responses"].values() if r["skipped"]]),
            "avg": avg,
            "pct": (sum(scored) / (5.0 * len(scored)) * 100) if scored else 0.0,
            "sum": sum(scored),
            "by_level": by_level,
        }

    def show_interview(self):
        self._clear()
        self._set_step(1)
        wrap = tk.Frame(self.body, bg=CANVAS)
        wrap.pack(fill="both", expand=True, padx=30, pady=24)

        self.rail = tk.Frame(wrap, bg=CANVAS, width=268)
        self.rail.pack(side="right", fill="y", padx=(22, 0))
        self.rail.pack_propagate(False)
        self._build_rail()

        self.stage = tk.Frame(wrap, bg=CANVAS)
        self.stage.pack(side="left", fill="both", expand=True)
        self._render_question()

    # -- right rail --------------------------------------------------------
    def _build_rail(self):
        outer, r = card(self.rail, pad=18)
        outer.pack(fill="both", expand=True)

        tk.Label(r, text="LIVE SCORE", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w")
        row = tk.Frame(r, bg=SURFACE)
        row.pack(anchor="w", pady=(2, 0))
        self.rail_avg = tk.Label(row, text="0.0", bg=SURFACE, fg=ACCENT,
                                 font=font(34, "bold", "display"))
        self.rail_avg.pack(side="left")
        tk.Label(row, text="/ 5", bg=SURFACE, fg=FAINT,
                 font=font(13, "bold")).pack(side="left", anchor="s", pady=(0, 9), padx=(5, 0))
        self.rail_sub = tk.Label(r, text="0 of 15 answered", bg=SURFACE, fg=MUTED, font=font(9))
        self.rail_sub.pack(anchor="w", pady=(0, 14))

        tk.Frame(r, bg=RULE, height=1).pack(fill="x", pady=(0, 14))
        tk.Label(r, text="PROGRESS", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 8))
        self.strip = tk.Canvas(r, bg=SURFACE, highlightthickness=0, bd=0, height=190, width=228)
        self.strip.pack(anchor="w")
        self.strip.bind("<Button-1>", self._strip_click)

        tk.Frame(r, bg=RULE, height=1).pack(fill="x", pady=14)
        tk.Label(r, text="BY DIFFICULTY", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 8))
        self.level_labels = {}
        for lvl in ("Basic", "Medium", "Advanced"):
            line = tk.Frame(r, bg=SURFACE)
            line.pack(fill="x", pady=2)
            tk.Label(line, text=lvl, bg=SURFACE, fg=INK, font=font(9)).pack(side="left")
            v = tk.Label(line, text="\u2013", bg=SURFACE, fg=MUTED, font=font(9, "bold", "mono"))
            v.pack(side="right")
            self.level_labels[lvl] = v

        tk.Frame(r, bg=SURFACE).pack(fill="both", expand=True)
        tk.Label(r, text="A skipped question counts as 0.", bg=SURFACE, fg=FAINT,
                 font=font(8), wraplength=200, justify="left").pack(anchor="w")

    CELL_W, CELL_H, GAP = 70, 32, 6

    def _strip_click(self, event):
        col = int(event.x // (self.CELL_W + self.GAP))
        row = int(event.y // (self.CELL_H + self.GAP))
        idx = row * 3 + col
        if 0 <= col < 3 and 0 <= idx < len(self.data["questions"]):
            self._save_current()
            self.q_index = idx
            self._render_question()

    def _draw_strip(self):
        self.strip.delete("all")
        for i, q in enumerate(self.data["questions"]):
            col, row = i % 3, i // 3
            x = col * (self.CELL_W + self.GAP)
            y = row * (self.CELL_H + self.GAP)
            r = self.data["responses"].get(i)
            if r and r["skipped"]:
                fill, fg, txt = SKIP_FILL, INK, "skip"
            elif r and r["score"] is not None:
                fill, fg, txt = SCORE_FILL[r["score"]], SCORE_TEXT[r["score"]], str(r["score"])
            else:
                fill, fg, txt = EMPTY_FILL, FAINT, "\u2013"
            border = ACCENT if i == self.q_index else RULE
            self.strip.create_rectangle(x, y, x + self.CELL_W, y + self.CELL_H,
                                        fill=fill, outline=border,
                                        width=2 if i == self.q_index else 1)
            self.strip.create_text(x + 9, y + self.CELL_H / 2, text="Q%d" % (i + 1),
                                   anchor="w", fill=fg, font=font(8, "bold", "mono"))
            self.strip.create_text(x + self.CELL_W - 9, y + self.CELL_H / 2, text=txt,
                                   anchor="e", fill=fg, font=font(9, "bold"))

    def _refresh_rail(self):
        s = self.stats()
        self.rail_avg.configure(text="%.1f" % s["avg"])
        self.rail_sub.configure(text="%d of %d answered  \u00b7  %d skipped"
                                     % (s["attempted"], s["total"], s["skipped"]))
        for lvl, v in self.level_labels.items():
            val = s["by_level"][lvl]
            v.configure(text="\u2013" if val is None else "%.1f" % val,
                        fg=MUTED if val is None else ACCENT)
        self._draw_strip()

    # -- question ----------------------------------------------------------
    def _render_question(self):
        for w in self.stage.winfo_children():
            w.destroy()
        q = self.data["questions"][self.q_index]
        r = self._response(self.q_index)

        outer, c = card(self.stage, pad=26)
        outer.pack(fill="both", expand=True)

        top = tk.Frame(c, bg=SURFACE)
        top.pack(fill="x")
        tk.Label(top, text="QUESTION %02d / %02d" % (self.q_index + 1, len(self.data["questions"])),
                 bg=SURFACE, fg=MUTED, font=font(9, "bold", "mono")).pack(side="left")
        badge = tk.Label(top, text=q["level"].upper(), bg="#EDF4F3", fg=ACCENT_DK,
                         font=font(8, "bold"), padx=9, pady=3)
        badge.pack(side="left", padx=12)
        tk.Label(top, text="%s  \u00b7  %s" % (self.data["topic"], q["qid"]), bg=SURFACE,
                 fg=FAINT, font=font(9, "normal", "mono")).pack(side="right")

        tk.Label(c, text=q["question"], bg=SURFACE, fg=INK, font=font(15, "bold", "display"),
                 wraplength=660, justify="left").pack(anchor="w", pady=(16, 16))

        exp = tk.Frame(c, bg="#F4F7F8")
        exp.pack(fill="x", pady=(0, 18))
        tk.Frame(exp, bg=ACCENT, width=3).pack(side="left", fill="y")
        inner = tk.Frame(exp, bg="#F4F7F8", padx=14, pady=12)
        inner.pack(side="left", fill="both", expand=True)
        tk.Label(inner, text="EXPECTED ANSWER  \u00b7  INTERVIEWER ONLY", bg="#F4F7F8",
                 fg=MUTED, font=font(8, "bold")).pack(anchor="w", pady=(0, 5))
        tk.Label(inner, text=q["expected"], bg="#F4F7F8", fg=INK, font=font(10),
                 wraplength=620, justify="left").pack(anchor="w")

        self.notes = TextBox(c, "What the candidate said", height=4,
                             hint="optional \u00b7 goes into the record")
        self.notes.pack(fill="x", pady=(0, 18))
        self.notes.set(r["notes"])

        tk.Label(c, text="SCORE THIS ANSWER", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 8))
        self.scorebar = ScoreBar(c, command=self._on_score)
        self.scorebar.pack(anchor="w")
        if r["score"] is not None and not r["skipped"]:
            self.scorebar.set(r["score"], fire=False)

        self.skip_note = tk.Label(c, text="", bg=SURFACE, fg=WARN, font=font(9, "bold"))
        self.skip_note.pack(anchor="w", pady=(10, 0))
        if r["skipped"]:
            self.skip_note.configure(text="Marked as skipped \u2014 counts as 0. "
                                          "Pick a score to undo.")

        tk.Frame(c, bg=SURFACE).pack(fill="both", expand=True)
        tk.Frame(c, bg=RULE, height=1).pack(fill="x", pady=(18, 16))

        nav = tk.Frame(c, bg=SURFACE)
        nav.pack(fill="x")
        if self.q_index > 0:
            Button(nav, "\u2190  Previous", self._prev, kind="quiet").pack(side="left")
        Button(nav, "Skip \u2014 didn't know", self._skip, kind="ghost").pack(side="left", padx=8)

        last = self.q_index == len(self.data["questions"]) - 1
        Button(nav, "Finish and review  \u2192" if last else "Save and next  \u2192",
               self._finish if last else self._next,
               kind="dark" if last else "primary").pack(side="right")

        self._refresh_rail()

    def _on_score(self, value):
        r = self._response(self.q_index)
        r["score"] = value
        r["skipped"] = False
        self.skip_note.configure(text="")
        self._refresh_rail()

    def _save_current(self):
        if hasattr(self, "notes") and self.notes.winfo_exists():
            self._response(self.q_index)["notes"] = self.notes.get()

    def _skip(self):
        self._save_current()
        r = self._response(self.q_index)
        r["skipped"] = True
        r["score"] = None
        self.scorebar.clear()
        self.skip_note.configure(text="Marked as skipped \u2014 counts as 0.")
        self._refresh_rail()
        if self.q_index < len(self.data["questions"]) - 1:
            self.after(220, self._next)

    def _next(self):
        self._save_current()
        r = self._response(self.q_index)
        if r["score"] is None and not r["skipped"]:
            messagebox.showinfo(
                "Score needed",
                "Give this answer a score from 1 to 5, or use Skip if the "
                "candidate didn't know it.")
            return
        if self.q_index < len(self.data["questions"]) - 1:
            self.q_index += 1
            self._render_question()

    def _prev(self):
        self._save_current()
        if self.q_index > 0:
            self.q_index -= 1
            self._render_question()

    def _finish(self):
        self._save_current()
        r = self._response(self.q_index)
        if r["score"] is None and not r["skipped"]:
            messagebox.showinfo("Score needed",
                                "Score the last answer, or skip it, before finishing.")
            return
        left = [i + 1 for i in range(len(self.data["questions"]))
                if i not in self.data["responses"]
                or (self.data["responses"][i]["score"] is None
                    and not self.data["responses"][i]["skipped"])]
        if left:
            go = messagebox.askyesno(
                "Unscored questions",
                "These questions have no score yet: %s.\n\nFinish anyway?"
                % ", ".join("Q%d" % n for n in left))
            if not go:
                return
        self.data["finished"] = datetime.now().isoformat(timespec="seconds")
        self.show_wrapup()

    # ==================================================================
    # 3. WRAP-UP
    # ==================================================================
    def show_wrapup(self):
        self._clear()
        self._set_step(2)
        scroll = Scroll(self.body)
        scroll.pack(fill="both", expand=True)
        page = tk.Frame(scroll.inner, bg=CANVAS, padx=34, pady=26)
        page.pack(fill="both", expand=True)
        s = self.stats()
        cand = self.data["candidate"]

        tk.Label(page, text="Wrap up the interview", bg=CANVAS, fg=INK,
                 font=font(20, "bold", "display")).pack(anchor="w")
        tk.Label(page, text="%s  \u00b7  %s  \u00b7  %s" % (
            cand.get("candidate", ""), self.data["topic"], self.data["interview_id"]),
                 bg=CANVAS, fg=MUTED, font=font(10)).pack(anchor="w", pady=(4, 20))

        metrics = tk.Frame(page, bg=CANVAS)
        metrics.pack(fill="x", pady=(0, 18))
        for label, value in (("Average score", "%.1f / 5" % s["avg"]),
                             ("Score", "%d%%" % round(s["pct"])),
                             ("Answered", "%d of %d" % (s["attempted"], s["total"])),
                             ("Skipped", str(s["skipped"]))):
            o, b = card(metrics, pad=16)
            o.pack(side="left", fill="x", expand=True, padx=(0, 12))
            tk.Label(b, text=label.upper(), bg=SURFACE, fg=MUTED,
                     font=font(8, "bold")).pack(anchor="w")
            tk.Label(b, text=value, bg=SURFACE, fg=INK,
                     font=font(20, "bold", "display")).pack(anchor="w", pady=(2, 0))

        # --- interviewer assessment ---
        o, c = card(page)
        o.pack(fill="x", pady=(0, 18))
        tk.Label(c, text="Your assessment", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w")
        tk.Label(c, text="Your own rating, independent of the calculated average.",
                 bg=SURFACE, fg=MUTED, font=font(9)).pack(anchor="w", pady=(3, 14))

        tk.Label(c, text="OVERALL INTERVIEWER RATING", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 8))
        self.rating_bar = ScoreBar(c)
        self.rating_bar.pack(anchor="w", pady=(0, 18))
        if self.data["rating"]:
            self.rating_bar.set(self.data["rating"], fire=False)

        tk.Label(c, text="RECOMMENDATION", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 8))
        pills = tk.Frame(c, bg=SURFACE)
        pills.pack(anchor="w", pady=(0, 18))
        self.decision_tiles = {}
        for d in DECISIONS:
            wrap = tk.Frame(pills, bg=RULE)
            wrap.pack(side="left", padx=(0, 8))
            lbl = tk.Label(wrap, text=d, bg=SURFACE, fg=INK, font=font(10, "bold"),
                           padx=18, pady=9, cursor="hand2")
            lbl.pack(padx=1, pady=1)
            lbl.bind("<Button-1>", lambda e, v=d: self._pick_decision(v))
            self.decision_tiles[d] = (wrap, lbl)
        if self.data["decision"]:
            self._pick_decision(self.data["decision"])

        self.comment_box = TextBox(c, "Comments", height=5, hint="optional")
        self.comment_box.pack(fill="x")
        self.comment_box.set(self.data["comment"])

        # --- review ---
        o2, rv = card(page)
        o2.pack(fill="x", pady=(0, 18))
        tk.Label(rv, text="Answer review", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w", pady=(0, 12))
        for i, q in enumerate(self.data["questions"]):
            r = self.data["responses"].get(i, {})
            line = tk.Frame(rv, bg=SURFACE)
            line.pack(fill="x", pady=1)
            tk.Label(line, text="Q%02d" % (i + 1), bg=SURFACE, fg=MUTED, width=4,
                     font=font(9, "bold", "mono")).pack(side="left")
            tk.Label(line, text=q["level"][:3].upper(), bg=SURFACE, fg=FAINT, width=5,
                     font=font(8, "bold")).pack(side="left")
            if r.get("skipped"):
                fill, fg, txt = SKIP_FILL, INK, "skip"
            elif r.get("score"):
                fill, fg, txt = SCORE_FILL[r["score"]], SCORE_TEXT[r["score"]], str(r["score"])
            else:
                fill, fg, txt = EMPTY_FILL, FAINT, "\u2013"
            tk.Label(line, text=txt, bg=fill, fg=fg, width=5, pady=3,
                     font=font(9, "bold")).pack(side="right", padx=(10, 0))
            tk.Label(line, text=q["question"], bg=SURFACE, fg=INK, font=font(9.5),
                     anchor="w", wraplength=760, justify="left").pack(side="left", fill="x")
            tk.Frame(rv, bg="#F0F3F5", height=1).pack(fill="x", pady=(4, 4))

        nav = tk.Frame(page, bg=CANVAS)
        nav.pack(fill="x")
        Button(nav, "\u2190  Back to questions", self._back_to_questions,
               kind="quiet").pack(side="left")
        Button(nav, "Continue to documents  \u2192", self._goto_export).pack(side="right")

    def _pick_decision(self, value):
        self.data["decision"] = value
        for d, (wrap, lbl) in self.decision_tiles.items():
            on = d == value
            wrap.configure(bg=HEADER if on else RULE)
            lbl.configure(bg=HEADER if on else SURFACE, fg="#FFFFFF" if on else INK)

    def _harvest_wrapup(self):
        self.data["rating"] = self.rating_bar.value
        self.data["comment"] = self.comment_box.get()

    def _back_to_questions(self):
        self._harvest_wrapup()
        self.show_interview()

    def _goto_export(self):
        self._harvest_wrapup()
        if not self.data["rating"]:
            messagebox.showinfo("Rating needed",
                                "Give an overall rating from 1 to 5 before continuing.")
            return
        self.show_export()

    # ==================================================================
    # 4. DOCUMENTS
    # ==================================================================
    def show_export(self):
        self._clear()
        self._set_step(3)
        scroll = Scroll(self.body)
        scroll.pack(fill="both", expand=True)
        page = tk.Frame(scroll.inner, bg=CANVAS, padx=34, pady=26)
        page.pack(fill="both", expand=True)

        tk.Label(page, text="Generate the record", bg=CANVAS, fg=INK,
                 font=font(20, "bold", "display")).pack(anchor="w")
        tk.Label(page, text="Everything is saved into one folder named after the interview ID.",
                 bg=CANVAS, fg=MUTED, font=font(10)).pack(anchor="w", pady=(4, 20))

        o, c = card(page)
        o.pack(fill="x", pady=(0, 18))
        tk.Label(c, text="Save location", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w", pady=(0, 12))
        row = tk.Frame(c, bg=SURFACE)
        row.pack(fill="x")
        box = tk.Frame(row, bg=RULE)
        box.pack(side="left", fill="x", expand=True)
        self.out_var = tk.StringVar(value=default_output_dir())
        tk.Entry(box, textvariable=self.out_var, bd=0, relief="flat", bg=SURFACE, fg=INK,
                 font=font(10), insertbackground=ACCENT).pack(fill="x", ipady=8, ipadx=9,
                                                              padx=1, pady=1)
        Button(row, "Browse", self._browse_out, kind="ghost").pack(side="left", padx=(10, 0))
        tk.Label(c, text="A folder named %s is created here, and every filename inside "
                         "starts with that ID." % self.data["interview_id"],
                 bg=SURFACE, fg=FAINT, font=font(9)).pack(anchor="w", pady=(8, 0))

        o2, t = card(page)
        o2.pack(fill="x", pady=(0, 18))
        tk.Label(t, text="Meeting transcript", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w")
        tk.Label(t, text="Optional. Attach the transcript file from the call \u2014 "
                         ".txt, .vtt, .csv or .docx.",
                 bg=SURFACE, fg=MUTED, font=font(9)).pack(anchor="w", pady=(3, 14))
        frow = tk.Frame(t, bg=SURFACE)
        frow.pack(fill="x")
        Button(frow, "Choose file", self._pick_transcript, kind="ghost").pack(side="left")
        self.file_lbl = tk.Label(frow, text="No file chosen", bg=SURFACE, fg=FAINT,
                                 font=font(9.5))
        self.file_lbl.pack(side="left", padx=12)
        self.anon_var = tk.IntVar(value=1)
        tk.Checkbutton(t, text="Anonymise names \u2014 replace the candidate and interviewer "
                               "with role labels and the interview ID",
                       variable=self.anon_var, bg=SURFACE, fg=INK, font=font(9.5),
                       activebackground=SURFACE, activeforeground=INK, selectcolor=SURFACE,
                       highlightthickness=0, bd=0, anchor="w").pack(anchor="w", pady=(12, 0))
        if self.data.get("transcript_source"):
            self.file_lbl.configure(text=os.path.basename(self.data["transcript_source"]),
                                    fg=INK)

        nav = tk.Frame(page, bg=CANVAS)
        nav.pack(fill="x")
        Button(nav, "\u2190  Back", self.show_wrapup, kind="quiet").pack(side="left")
        Button(nav, "Generate documents", self._generate, kind="dark").pack(side="right")

    def _browse_out(self):
        d = filedialog.askdirectory(title="Choose where to save interview records")
        if d:
            self.out_var.set(d)

    def _pick_transcript(self):
        p = filedialog.askopenfilename(
            title="Choose the meeting transcript",
            filetypes=[("Transcript files", "*.txt *.vtt *.csv *.docx *.md"),
                       ("All files", "*.*")])
        if p:
            self.data["transcript_source"] = p
            self.file_lbl.configure(text=os.path.basename(p), fg=INK)

    def _generate(self):
        base = self.out_var.get().strip() or default_output_dir()
        folder = os.path.join(base, self.data["interview_id"])
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Cannot save here",
                                 "That folder could not be created:\n%s" % exc)
            return

        prefix = "%s_%s" % (self.data["interview_id"],
                            safe_name(self.data["candidate"].get("candidate")))
        made = []
        try:
            p = os.path.join(folder, prefix + "_Transcript.docx")
            self._build_interview_transcript().save(p)
            made.append(p)

            src = self.data.get("transcript_source")
            if src and os.path.isfile(src):
                p2 = os.path.join(folder, prefix + "_MeetingTranscript.docx")
                self._build_meeting_transcript(src).save(p2)
                made.append(p2)
        except Exception as exc:
            messagebox.showerror("Could not write the documents", str(exc))
            return
        self.show_done(folder, made)

    def show_done(self, folder, files):
        self._clear()
        page = tk.Frame(self.body, bg=CANVAS, padx=34, pady=40)
        page.pack(fill="both", expand=True)
        o, c = card(page, pad=30)
        o.pack(fill="x")
        tk.Label(c, text="Documents saved", bg=SURFACE, fg=ACCENT,
                 font=font(22, "bold", "display")).pack(anchor="w")
        tk.Label(c, text=folder, bg=SURFACE, fg=MUTED, font=font(10, "normal", "mono"),
                 wraplength=820, justify="left").pack(anchor="w", pady=(6, 18))
        for f in files:
            tk.Label(c, text="\u2022   " + os.path.basename(f), bg=SURFACE, fg=INK,
                     font=font(10)).pack(anchor="w", pady=1)
        row = tk.Frame(c, bg=SURFACE)
        row.pack(anchor="w", pady=(24, 0))
        Button(row, "Open folder", lambda: open_folder(folder), kind="primary").pack(side="left")
        Button(row, "New interview", self._new_interview,
               kind="ghost").pack(side="left", padx=10)

    def _new_interview(self):
        if not messagebox.askyesno("Start a new interview",
                                   "Clear this interview and start a new one?"):
            return
        self.reset_data()
        self._set_id_chip()
        self.show_setup()

    # ==================================================================
    # Documents
    # ==================================================================
    def _duration(self):
        try:
            a = datetime.fromisoformat(self.data["started"])
            b = datetime.fromisoformat(self.data["finished"])
            return "%d min" % int((b - a).total_seconds() // 60)
        except Exception:
            return "Not recorded"

    def _build_interview_transcript(self):
        """Plain text record of the interview - this is what the agent reads."""
        d = Docx()
        cand = self.data["candidate"]
        s = self.stats()

        def L(text="", bold=False):
            d.para(text, bold=bold, size=11, space_after=0)

        L("INTERVIEW TRANSCRIPT", bold=True)
        L()
        L("Interview ID: %s" % self.data["interview_id"])
        L("Candidate: %s" % (cand.get("candidate") or "Not recorded"))
        L("Candidate ID: %s" % (cand.get("cid") or "Not recorded"))
        L("Requisition ID: %s" % (cand.get("req") or "Not recorded"))
        L("Team: %s" % (cand.get("team") or "Not recorded"))
        L("Role: %s" % (cand.get("role") or "Not recorded"))
        L("Grade: %s" % (cand.get("grade") or "Not recorded"))
        L("Topic: %s" % self.data["topic"])
        L("Interview date: %s" % (cand.get("date") or "Not recorded"))
        L("Interview time: %s" % (cand.get("time") or "Not recorded"))
        L("Interviewer: %s" % (cand.get("interviewer") or "Not recorded"))
        L("Duration: %s" % self._duration())
        L("Question bank: %s" % (self.data.get("source_note") or "Not recorded"))
        L()

        L("SUMMARY", bold=True)
        L()
        L("Questions asked: %d" % s["total"])
        L("Answered: %d" % (s["attempted"] - s["skipped"]))
        L("Skipped: %d" % s["skipped"])
        L("Total marks: %d out of %d" % (s["sum"], 5 * s["total"]))
        L("Average mark: %.1f out of 5" % s["avg"])
        L("Score: %d%%" % round(s["pct"]))
        for lvl in ("Basic", "Medium", "Advanced"):
            val = s["by_level"][lvl]
            L("%s average: %s" % (lvl, "Not attempted" if val is None else "%.1f out of 5" % val))
        L()

        L("QUESTIONS AND ANSWERS", bold=True)
        L()
        for i, q in enumerate(self.data["questions"]):
            r = self.data["responses"].get(i, {})
            L("Q%d (%s)" % (i + 1, q["level"]), bold=True)
            L("Question: %s" % q["question"])
            L("Expected answer: %s" % q["expected"])
            if r.get("skipped"):
                L("Candidate answer: Skipped - the candidate did not know the answer.")
                L("Mark: 0 out of 5")
            else:
                L("Candidate answer: %s" % (r.get("notes") or "Not recorded"))
                L("Mark: %s" % ("Not scored" if r.get("score") is None
                                else "%d out of 5" % r["score"]))
            L()

        L("FINAL ASSESSMENT", bold=True)
        L()
        L("Average mark from the question scores: %.1f out of 5 (%d%%)"
          % (s["avg"], round(s["pct"])))
        L("Interviewer rating: %s out of 5" % (self.data["rating"] or "Not recorded"))
        L("Recommendation: %s" % (self.data["decision"] or "Not recorded"))
        L("Interviewer comment: %s" % (self.data["comment"] or "No comment given."))
        L()
        L("End of transcript.")
        return d

    def _build_meeting_transcript(self, src):
        """Plain text copy of the uploaded meeting transcript."""
        text = read_text_file(src)
        hits = 0
        anon = bool(self.anon_var.get()) if hasattr(self, "anon_var") else True
        if anon:
            cand = self.data["candidate"]
            text, hits = anonymize(text, {
                cand.get("candidate", ""): "Candidate",
                cand.get("interviewer", ""): "Interviewer",
            })
        self.data["anonymised"] = anon

        d = Docx()

        def L(line="", bold=False):
            d.para(line, bold=bold, size=11, space_after=0)

        L("MEETING TRANSCRIPT", bold=True)
        L()
        L("Interview ID: %s" % self.data["interview_id"])
        L("Topic: %s" % self.data["topic"])
        L("Interview date: %s" % (self.data["candidate"].get("date") or "Not recorded"))
        L("Source file: %s" % os.path.basename(src))
        if anon:
            L("Anonymised: Yes - %d name reference%s replaced with Interviewer and Candidate"
              % (hits, "" if hits == 1 else "s"))
        else:
            L("Anonymised: No")
        L()
        L("TRANSCRIPT", bold=True)
        L()
        blank = 0
        for line in text.splitlines():
            line = line.rstrip()
            if not line:
                blank += 1
                if blank <= 1:
                    L()
                continue
            blank = 0
            L(line)
        L()
        L("End of transcript.")
        return d


# --------------------------------------------------------------------------

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
