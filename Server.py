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
import json
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
    "VBA": _bank("VBA", [
        ("What is VBA and where does it run?",
         "Visual Basic for Applications - the macro language built into Office. It runs inside the host application through the Visual Basic Editor and drives that application's object model."),
        ("What is the difference between a Sub and a Function?",
         "A Sub performs actions and returns nothing. A Function returns a value and, if public in a standard module, can be called from a worksheet formula."),
        ("How do you declare a variable, and what does Option Explicit do?",
         "Dim name As Type declares a variable. Option Explicit at the top of a module forces every variable to be declared, so typos are caught at compile time instead of silently creating Variants."),
        ("What is the difference between Range and Cells?",
         "Range refers to one or more cells and accepts addresses and named ranges. Cells(row, column) addresses a single cell by index, which is convenient inside loops. Range is more flexible; Cells is easier to compute."),
        ("How do you record a macro, and why is recorded code usually poor?",
         "Developer tab, Record Macro. The recorder uses Select and Activate, assumes the active sheet, hardcodes what was clicked, and captures incidental actions, so the result is slow and breaks when anything moves."),
        ("Why should you avoid Select and Activate, and what do you write instead?",
         "They depend on what happens to be active and cost a screen redraw each time. Use direct references instead, for example ws.Range(\"A1\").Value = 1, which is faster and safe when other workbooks are open."),
        ("Explain For Each versus For Next.",
         "For Each walks a collection without an index and reads cleanly for ranges and worksheets. For Next uses a counter, which you need when the index matters or when stepping backwards - for example deleting rows from the bottom up."),
        ("How do you handle errors in VBA?",
         "On Error GoTo label with a handler that reports Err.Number and Err.Description, then Resume or exits cleanly. On Error Resume Next only around a statement that is expected to fail, cleared straight away with On Error GoTo 0."),
        ("What are workbook and worksheet events? Give examples.",
         "Procedures that fire on triggers: Workbook_Open, Workbook_BeforeSave, Worksheet_Change, Worksheet_SelectionChange. They live in the workbook or sheet module, not a standard module, and often need EnableEvents toggled to avoid recursion."),
        ("How do you find the last used row reliably?",
         "Cells(Rows.Count, 1).End(xlUp).Row from the bottom up, or Find with SearchDirection:=xlPrevious for the whole sheet. UsedRange alone is unreliable because it goes stale after deletions."),
        ("How would you speed up a slow macro?",
         "Turn off ScreenUpdating, set Calculation to manual, disable events, stop using Select, read the range into a Variant array, process in memory, and write back in one assignment. Restore every setting in the error handler as well as at the end."),
        ("Compare arrays and ranges for bulk processing.",
         "Reading Range.Value into a two-dimensional Variant array is a single call across the COM boundary; looping cell by cell is thousands of calls. Process in the array and assign back to a range of the same size."),
        ("What is early versus late binding, and what are the trade-offs?",
         "Early binding sets a library reference at design time and uses New, giving IntelliSense, type checking and speed, but it breaks if the target machine has a different version. Late binding uses CreateObject with Object variables - portable across versions, no IntelliSense, slightly slower."),
        ("How do you structure and secure VBA in a workbook that will be shared?",
         "Split code into modules by purpose, keep configuration in constants or a settings sheet, avoid hardcoded paths, add a version note and changelog. Lock the project with a password (weak, deters casual edits only) and sign it with a digital certificate so it runs under macro security."),
        ("A macro works on your machine but fails on a colleague's. How do you debug it?",
         "Check missing references, Office version and bitness, regional settings for dates and decimal separators, file paths and mapped drives, trusted locations and macro security, conflicting add-ins, and screen or resolution assumptions. Add logging and Debug.Print, and replace assumptions with defensive checks."),
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

TOPICS = ["SQL", "VBA", "Excel", "Python", "AI"]
LEVELS = ["Basic", "Medium", "Advanced"]
LEVEL_WEIGHT = {"Basic": 1, "Medium": 3, "Advanced": 5}
PER_LEVEL = 3   # questions asked per difficulty band
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
# Minimal PDF writer (no external packages)
# Helvetica and Helvetica-Bold are built into every PDF reader, so nothing
# has to be embedded. The width tables below let text be wrapped accurately.
# --------------------------------------------------------------------------

_W_REG = ("278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 "
          "556 556 556 556 556 556 556 556 556 556 278 278 584 584 584 556 "
          "1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 "
          "667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 "
          "333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 "
          "556 556 333 500 278 556 500 722 500 500 500 334 260 334 584")

_W_BLD = ("278 333 474 556 556 889 722 238 333 333 389 584 278 333 278 278 "
          "556 556 556 556 556 556 556 556 556 556 333 333 584 584 584 611 "
          "975 722 722 722 722 667 611 778 722 278 556 722 611 833 722 778 "
          "667 778 722 667 611 722 667 944 667 667 611 333 278 333 584 556 "
          "333 556 611 556 611 556 333 611 611 278 278 556 278 889 611 611 "
          "611 611 389 556 333 611 556 778 556 556 500 389 280 389 584")


def _widths(spec):
    return {chr(32 + i): int(v) for i, v in enumerate(spec.split())}


WIDTH_REG = _widths(_W_REG)
WIDTH_BLD = _widths(_W_BLD)

# Characters a PDF's WinAnsi encoding will not take, folded to plain ASCII.
_FOLD = {
    "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
    "\u2013": "-", "\u2014": "-", "\u2022": "-", "\u00b7": "-",
    "\u2026": "...", "\u00a0": " ", "\u2192": "->", "\u2190": "<-",
    "\u2265": ">=", "\u2264": "<=", "\u00d7": "x", "\u2013": "-",
}


def pdf_text(value):
    out = []
    for ch in (value or ""):
        if ch in _FOLD:
            out.append(_FOLD[ch])
        elif ord(ch) < 32:
            out.append(" ")
        elif ord(ch) < 127:
            out.append(ch)
        else:
            try:
                ch.encode("latin-1")
                out.append(ch)
            except Exception:
                out.append("?")
    return "".join(out)


def pesc(value):
    return (pdf_text(value).replace("\\", r"\\")
            .replace("(", r"\(").replace(")", r"\)"))


class Pdf(object):
    """A4 page builder with a simple top-down text cursor."""

    W, H = 595.28, 841.89

    def __init__(self, margin=52, footer=""):
        self.margin = margin
        self.footer = footer
        self.pages = []
        self.ops = None
        self.y = 0.0
        self.on_new_page = None
        self.new_page()

    # -- geometry ----------------------------------------------------------
    @property
    def right(self):
        return self.W - self.margin

    @property
    def usable(self):
        return self.W - 2 * self.margin

    def new_page(self):
        self.ops = []
        self.pages.append(self.ops)
        self.y = self.H - self.margin
        if self.on_new_page:
            self.on_new_page(self)

    def need(self, height):
        if self.y - height < self.margin + 34:
            self.new_page()
            return True
        return False

    def space(self, amount):
        self.y -= amount

    # -- primitives --------------------------------------------------------
    def _col(self, rgb):
        r, g, b = rgb
        return "%.3f %.3f %.3f" % (r / 255.0, g / 255.0, b / 255.0)

    def text(self, x, y, value, size=10, bold=False, color=(26, 35, 43)):
        self.ops.append("BT /%s %.2f Tf %s rg %.2f %.2f Td (%s) Tj ET"
                        % ("F2" if bold else "F1", size, self._col(color),
                           x, y, pesc(value)))

    def rect(self, x, y, w, h, color):
        self.ops.append("%s rg %.2f %.2f %.2f %.2f re f"
                        % (self._col(color), x, y, w, h))

    def line(self, x1, y1, x2, y2, color=(224, 230, 234), width=0.8):
        self.ops.append("%s RG %.2f w %.2f %.2f m %.2f %.2f l S"
                        % (self._col(color), width, x1, y1, x2, y2))

    # -- measuring and wrapping -------------------------------------------
    def width(self, value, size, bold=False):
        table = WIDTH_BLD if bold else WIDTH_REG
        return sum(table.get(c, 556) for c in pdf_text(value)) * size / 1000.0

    def wrap(self, value, size, max_width, bold=False):
        lines = []
        for chunk in pdf_text(value).split("\n"):
            words, line = chunk.split(), ""
            for word in words:
                trial = (line + " " + word).strip()
                if self.width(trial, size, bold) <= max_width or not line:
                    line = trial
                else:
                    lines.append(line)
                    line = word
            lines.append(line)
        return lines or [""]

    # -- flowing content ---------------------------------------------------
    def para(self, value, size=9.5, bold=False, color=(60, 72, 82),
             indent=0, lead=1.45, after=0, width=None):
        max_w = width if width is not None else self.usable - indent
        step = size * lead
        for line in self.wrap(value, size, max_w, bold):
            self.need(step)
            self.y -= step
            self.text(self.margin + indent, self.y, line, size, bold, color)
        self.y -= after
        return self.y

    def heading(self, value, size=11.5, color=(26, 35, 43), rule=True,
                before=20, after=10):
        self.need(before + size + 14)
        self.y -= before
        self.y -= size
        self.text(self.margin, self.y, value, size, True, color)
        if rule:
            self.y -= 6
            self.line(self.margin, self.y, self.right, self.y)
        self.y -= after
        return self.y

    def save(self, path):
        pages, total = self.pages, len(self.pages)
        for n, ops in enumerate(pages, start=1):
            note = self.footer
            self.ops = ops
            self.line(self.margin, self.margin + 20, self.right,
                      self.margin + 20, (232, 237, 240), 0.6)
            self.text(self.margin, self.margin + 8, note, 7.5, False, (150, 160, 168))
            label = "Page %d of %d" % (n, total)
            self.text(self.right - self.width(label, 7.5), self.margin + 8,
                      label, 7.5, False, (150, 160, 168))

        objs = []

        def add(body):
            objs.append(body)
            return len(objs)

        kids_id = 2
        add("<< /Type /Catalog /Pages 2 0 R >>")            # 1
        add("")                                             # 2 placeholder
        f1 = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica "
                 "/Encoding /WinAnsiEncoding >>")
        f2 = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold "
                 "/Encoding /WinAnsiEncoding >>")

        page_ids, streams = [], []
        for ops in pages:
            content = "\n".join(ops).encode("latin-1", "replace")
            sid = add("<< /Length %d >>\nstream\n%s\nendstream"
                      % (len(content), content.decode("latin-1")))
            pid = add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %.2f %.2f] "
                      "/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> >> "
                      "/Contents %d 0 R >>" % (self.W, self.H, f1, f2, sid))
            streams.append(sid)
            page_ids.append(pid)

        objs[kids_id - 1] = ("<< /Type /Pages /Count %d /Kids [%s] >>"
                             % (len(page_ids),
                                " ".join("%d 0 R" % i for i in page_ids)))

        out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for i, body in enumerate(objs, start=1):
            offsets.append(len(out))
            out += ("%d 0 obj\n%s\nendobj\n" % (i, body)).encode("latin-1", "replace")
        xref = len(out)
        out += ("xref\n0 %d\n" % (len(objs) + 1)).encode()
        out += b"0000000000 65535 f \n"
        for off in offsets[1:]:
            out += ("%010d 00000 n \n" % off).encode()
        out += ("trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n"
                % (len(objs) + 1, xref)).encode()
        with open(path, "wb") as f:
            f.write(bytes(out))


# --------------------------------------------------------------------------
# Question rotation - each candidate gets a different slice of the bank
# --------------------------------------------------------------------------

def rotation_file():
    try:
        home = os.path.expanduser("~")
    except Exception:
        home = "."
    return os.path.join(home, ".interview_assistant_rotation.json")


def read_rotation():
    try:
        with open(rotation_file(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def bump_rotation(topics):
    """Move each topic on by one so the next candidate starts a step later."""
    data = read_rotation()
    for t in topics:
        try:
            data[t] = int(data.get(t, 0)) + 1
        except Exception:
            data[t] = 1
    try:
        with open(rotation_file(), "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass
    return data


def rotation_for(topic):
    try:
        return int(read_rotation().get(topic, 0))
    except Exception:
        return 0


def select_rotated(questions, topic, per_level=PER_LEVEL):
    """Pick `per_level` questions from each difficulty band, cycling per candidate.

    Candidate 1 gets questions 1,2,3 of a band, candidate 2 gets 2,3,4,
    candidate 3 gets 3,4,5, then it wraps back round to 4,5,1.
    """
    k = rotation_for(topic)
    out = []
    for level in LEVELS:
        band = [q for q in questions if q.get("level") == level]
        if not band:
            continue
        n = len(band)
        take = min(per_level, n)
        start = k % n
        for step in range(take):
            out.append(dict(band[(start + step) % n]))
    if not out:
        out = [dict(q) for q in questions]
    for i, q in enumerate(out, start=1):
        q["no"] = i
    return out


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
    "ID,Topic,Level,Question,ExpectedAnswer\r\n"
    "1,SQL,Basic,What does SELECT DISTINCT do?,"
    "Returns only unique rows for the selected columns.\r\n"
    "2,SQL,Medium,UNION vs UNION ALL?,"
    "UNION removes duplicates and sorts; UNION ALL keeps everything and is faster.\r\n"
    "3,SQL,Advanced,How do you tune a slow query?,"
    "Read the execution plan; check indexes; statistics; non-SARGable predicates.\r\n"
)


def norm_key(text):
    return re.sub(r"[^a-z0-9]", "", (text or "").lower())


def filename_matches_topic(path, topic):
    """The upload has to be named after the topic - python_bank.xlsx for Python."""
    base = norm_key(os.path.splitext(os.path.basename(path))[0])
    want = norm_key(topic)
    return bool(want) and base.startswith(want)


def topic_matches(cell, wanted):
    a, b = norm_key(cell), norm_key(wanted)
    if not a or not b:
        return False
    return a == b or (len(b) >= 2 and b in a) or (len(a) >= 3 and a in b)


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
            "No Question column found.\n\nThe first row must be a header, in the "
            "order ID, Topic, Level, Question, ExpectedAnswer.\n\nColumns found: %s"
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

    if not raw:
        raise ValueError("No question rows were found in that file.")
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
        sub = q["topic"] if q["topic"] and not topic_matches(q["topic"], topic) else ""
        out.append({
            "qid": q["qid"] or "%s-%02d" % (norm_key(topic)[:3].upper() or "QST", i),
            "no": i,
            "level": q["level"],
            "question": q["question"],
            "expected": q["expected"],
            "subtopic": sub,
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
            lbl = tk.Label(wrap, text=str(n), bg=SURFACE, fg=INK, width=4,
                           font=font(11, "bold", "display"), cursor="hand2", pady=2)
            lbl.pack(padx=1, pady=1)
            lbl.bind("<Button-1>", lambda e, v=n: self.set(v))
            lbl.bind("<Enter>", lambda e, v=n: self._hover(v, True))
            lbl.bind("<Leave>", lambda e, v=n: self._hover(v, False))
            self.cells[n] = (wrap, lbl)
        tk.Label(self, text="1  poor      3  adequate      5  excellent", bg=bg,
                 fg=FAINT, font=font(8)).pack(anchor="w", pady=(4, 0))

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


def hex_rgb(value):
    value = (value or "#000000").lstrip("#")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


INK_RGB = hex_rgb(INK)
MUTED_RGB = hex_rgb(MUTED)
FAINT_RGB = hex_rgb(FAINT)
ACCENT_RGB = hex_rgb(ACCENT)
HEADER_RGB = hex_rgb(HEADER)
PANEL_RGB = (245, 248, 249)
RULE_RGB = (224, 231, 235)


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
            "topics": [],       # chosen topics, in click order
            "banks": {},        # topic -> {"questions": [...], "note": str}
            "responses": {},    # topic -> {question index: response}
            "rating": None,
            "comment": "",
            "decision": None,
            "started": None,
            "finished": None,
            "transcript_source": None,
            "anonymised": False,
            "include_answers": True,
            "ended_early": False,
        }
        self.sec = 0
        self.q_index = 0

    # -- current section shortcuts ----------------------------------------
    @property
    def topic(self):
        tops = self.data["topics"]
        return tops[self.sec] if 0 <= self.sec < len(tops) else None

    def questions_for(self, topic):
        return self.data["banks"].get(topic, {}).get("questions", [])

    def responses_for(self, topic):
        return self.data["responses"].setdefault(topic, {})

    @property
    def qs(self):
        return self.questions_for(self.topic)

    @property
    def rs(self):
        return self.responses_for(self.topic)

    def asked_topics(self):
        """Topics that actually got at least one scored or skipped answer."""
        out = []
        for t in self.data["topics"]:
            rs = self.data["responses"].get(t, {})
            if any(r.get("skipped") or r.get("score") is not None for r in rs.values()):
                out.append(t)
        return out

    # -- chrome ------------------------------------------------------------
    def _build_chrome(self):
        bar = tk.Frame(self, bg=HEADER, height=64)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        left = tk.Frame(bar, bg=HEADER)
        left.pack(side="left", padx=26)
        tk.Label(left, text=APP_NAME, bg=HEADER, fg="#FFFFFF",
                 font=font(14, "bold", "display")).pack(side="left", pady=16)
        self.id_chip = tk.Label(left, text="", bg=HEADER, fg=ACCENT_LT,
                                font=font(10, "bold", "mono"))
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
        tk.Label(page, text="Enter the candidate's details, pick the topics in the order "
                            "you want to run them, and load the questions for each.",
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
        for k, v in (self.data.get("candidate") or {}).items():
            if k in self.fields and v:
                self.fields[k].set(v)

        tk.Label(c, text="Candidate name and interviewer name are required. Everything else is optional.",
                 bg=SURFACE, fg=FAINT, font=font(9)).pack(anchor="w")

        # --- topics ---
        outer2, t = card(page)
        outer2.pack(fill="x", pady=(0, 18))
        tk.Label(t, text="Topics", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w")
        tk.Label(t, text="Click one or more. The number on each tile is the order they "
                         "will be asked in \u2014 click again to remove. Each topic runs "
                         "%d questions: %d basic, %d medium, %d advanced."
                         % (PER_LEVEL * 3, PER_LEVEL, PER_LEVEL, PER_LEVEL),
                 bg=SURFACE, fg=MUTED, font=font(9.5), wraplength=860,
                 justify="left").pack(anchor="w", pady=(3, 14))

        tiles = tk.Frame(t, bg=SURFACE)
        tiles.pack(anchor="w")
        self.topic_tiles = {}
        for topic in TOPICS:
            wrap = tk.Frame(tiles, bg=RULE)
            wrap.pack(side="left", padx=(0, 10))
            lbl = tk.Label(wrap, text=topic, bg=SURFACE, fg=INK, font=font(11, "bold"),
                           padx=18, pady=13, cursor="hand2", width=10)
            lbl.pack(padx=1, pady=1)
            lbl.bind("<Button-1>", lambda e, tp=topic: self._toggle_topic(tp))
            self.topic_tiles[topic] = (wrap, lbl)

        tk.Frame(t, bg=RULE, height=1).pack(fill="x", pady=(18, 0))
        self.tstack = tk.Frame(t, bg=SURFACE)
        self.tstack.pack(fill="x")

        srow = tk.Frame(t, bg=SURFACE)
        srow.pack(fill="x", pady=(14, 0))
        Button(srow, "Save a blank template", self._save_template,
               kind="ghost", size=9, pady=10).pack(side="left")
        tk.Label(srow, text="Uploads must be CSV or Excel, named after the topic \u2014 "
                            "python_questions.xlsx for Python.",
                 bg=SURFACE, fg=FAINT, font=font(9)).pack(side="left", padx=12)

        # --- counts ---
        self.counts = tk.Frame(page, bg=CANVAS)
        self.counts.pack(fill="x", pady=(0, 18))
        self.count_labels = {}
        for name in ("Topics", "Questions", "Basic", "Medium", "Advanced"):
            box_o, box = card(self.counts, pad=13, bg="#F5F8F9", border="#E6ECEE")
            box_o.pack(side="left", padx=(0, 10))
            tk.Label(box, text=name.upper(), bg="#F5F8F9", fg=MUTED,
                     font=font(8, "bold")).pack()
            v = tk.Label(box, text="0", bg="#F5F8F9", fg=ACCENT,
                         font=font(19, "bold", "display"), width=7)
            v.pack()
            self.count_labels[name] = v

        row = tk.Frame(page, bg=CANVAS)
        row.pack(fill="x")
        self.btn_start = Button(row, "Start interview  \u2192", self._start_interview)
        self.btn_start.pack(side="left")

        self._paint_tiles()
        self._render_topic_rows()

    # -- topic selection ---------------------------------------------------
    def _toggle_topic(self, topic):
        tops = self.data["topics"]
        if topic in tops:
            tops.remove(topic)
            self.data["banks"].pop(topic, None)
            self.data["responses"].pop(topic, None)
        else:
            tops.append(topic)
        self._paint_tiles()
        self._render_topic_rows()

    def _paint_tiles(self):
        tops = self.data["topics"]
        for name, (wrap, lbl) in self.topic_tiles.items():
            if name in tops:
                n = tops.index(name) + 1
                wrap.configure(bg=ACCENT)
                lbl.configure(bg=ACCENT, fg="#FFFFFF", text="%d.  %s" % (n, name))
            else:
                wrap.configure(bg=RULE)
                lbl.configure(bg=SURFACE, fg=INK, text=name)

    def _render_topic_rows(self):
        for w in self.tstack.winfo_children():
            w.destroy()
        tops = self.data["topics"]
        if not tops:
            tk.Label(self.tstack, text="No topics chosen yet.", bg=SURFACE, fg=FAINT,
                     font=font(9.5)).pack(anchor="w", pady=16)
            self._refresh_counts()
            return

        for i, topic in enumerate(tops):
            line = tk.Frame(self.tstack, bg=SURFACE)
            line.pack(fill="x", pady=(14, 0))
            tk.Label(line, text=str(i + 1), bg=ACCENT, fg="#FFFFFF", width=3,
                     font=font(10, "bold")).pack(side="left", pady=2)
            tk.Label(line, text=topic, bg=SURFACE, fg=INK, font=font(11, "bold"),
                     width=11, anchor="w").pack(side="left", padx=12)
            Button(line, "Built-in", lambda tp=topic: self._load_questions(tp),
                   kind="ghost", size=9, pady=8, padx=14).pack(side="left")
            Button(line, "Upload CSV / Excel", lambda tp=topic: self._upload_questions(tp),
                   kind="ghost", size=9, pady=8, padx=14).pack(side="left", padx=8)
            bank = self.data["banks"].get(topic)
            note = bank["note"] if bank else "No questions loaded."
            tk.Label(self.tstack, text=note, bg=SURFACE,
                     fg=INK if bank else FAINT, font=font(9), anchor="w",
                     justify="left", wraplength=860).pack(anchor="w", padx=(52, 0),
                                                          pady=(4, 0))
        self._refresh_counts()

    def _refresh_counts(self):
        tops = self.data["topics"]
        allq = [q for t in tops for q in self.questions_for(t)]
        self.count_labels["Topics"].configure(text=str(len(tops)))
        self.count_labels["Questions"].configure(text=str(len(allq)))
        for lvl in LEVELS:
            self.count_labels[lvl].configure(
                text=str(len([q for q in allq if q["level"] == lvl])))
        ready = bool(tops) and all(self.data["banks"].get(t) for t in tops)
        if hasattr(self, "btn_start"):
            self.btn_start.set_enabled(ready)

    # -- loading questions -------------------------------------------------
    def _collect_candidate(self):
        return {k: f.get() for k, f in self.fields.items()}

    def _ready_to_load(self):
        cand = self._collect_candidate()
        if not cand["candidate"] or not cand["interviewer"]:
            messagebox.showinfo("Details needed",
                                "Enter the candidate name and the interviewer name first.")
            return False
        self.data["candidate"] = cand
        if not self.data["interview_id"]:
            self.data["interview_id"] = make_interview_id(cand["candidate"], cand["cid"])
        self._set_id_chip()
        return True

    def _describe(self, picked, pool, topic, source):
        counts = ", ".join("%d %s" % (len([q for q in picked if q["level"] == l]), l.lower())
                           for l in LEVELS)
        return ("%s \u2014 %d of %d questions selected (%s). Rotation set %d, so the "
                "next candidate gets a different slice."
                % (source, len(picked), pool, counts, rotation_for(topic) + 1))

    def _apply_bank(self, topic, picked, note, pool=None, source="Built-in bank",
                    path=None):
        self.data["banks"][topic] = {
            "questions": picked,
            "note": note,
            "pool": pool if pool is not None else [dict(q) for q in picked],
            "source": source,
            "path": path,
        }
        self.data["responses"][topic] = {}
        if hasattr(self, "tstack") and self.tstack.winfo_exists():
            self._render_topic_rows()

    def _rebuild_bank(self, topic, bank):
        """Reuse a previous candidate's question source, one rotation step on."""
        pool = bank.get("pool") or []
        if not pool:
            return
        picked = select_rotated(pool, topic)
        source = bank.get("source") or "Built-in bank"
        note = self._describe(picked, len(pool), topic, source)
        if bank.get("path"):
            note += "  Carried over from the last candidate."
        self._apply_bank(topic, picked, note, pool=pool, source=source,
                         path=bank.get("path"))

    def _load_questions(self, topic):
        if not self._ready_to_load():
            return
        pool = [dict(q) for q in QUESTIONS[topic]]
        picked = select_rotated(pool, topic)
        self._apply_bank(topic, picked,
                         self._describe(picked, len(pool), topic, "Built-in bank"),
                         pool=pool, source="Built-in bank")

    def _upload_questions(self, topic):
        if not self._ready_to_load():
            return
        path = filedialog.askopenfilename(
            title="Choose the %s question bank" % topic,
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
        if not filename_matches_topic(path, topic):
            messagebox.showwarning(
                "File name doesn't match the topic",
                "The file for %s has to be named starting with \"%s\".\n\n"
                "You chose: %s\n\nRename it to something like %s_questions.xlsx "
                "and try again."
                % (topic, topic, os.path.basename(path),
                   norm_key(topic) or topic.lower()))
            return
        try:
            questions, _ = load_questions_from_file(path, topic)
        except ValueError as exc:
            messagebox.showwarning("Could not read that file", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Could not read that file",
                                 "%s\n\nCheck the file opens in Excel and try again." % exc)
            return
        if not questions:
            messagebox.showwarning("No questions found",
                                   "No usable question rows were found in that file.")
            return
        picked = select_rotated(questions, topic)
        self._apply_bank(topic, picked,
                         self._describe(picked, len(questions), topic,
                                        os.path.basename(path)),
                         pool=questions, source=os.path.basename(path), path=path)

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
            "Columns, in this order: ID, Topic, Level, Question, ExpectedAnswer.\n\n"
            "Question is the only one that must be filled in. Level accepts Basic, "
            "Medium or Advanced and decides which band a question sits in.\n\n"
            "Name the file after the topic it belongs to \u2014 python_questions.xlsx, "
            "sql_bank.csv, vba_set1.xlsx \u2014 or the app will not accept it.")

    def _start_interview(self):
        cand = self._collect_candidate()
        if not cand["candidate"] or not cand["interviewer"]:
            messagebox.showinfo("Details needed",
                                "Enter the candidate name and the interviewer name first.")
            return
        self.data["candidate"] = cand
        if not self.data["interview_id"]:
            self.data["interview_id"] = make_interview_id(cand["candidate"], cand["cid"])
            self._set_id_chip()
        if not self.data["topics"]:
            messagebox.showinfo("Pick a topic", "Choose at least one topic first.")
            return
        missing = [t for t in self.data["topics"] if not self.data["banks"].get(t)]
        if missing:
            messagebox.showinfo("Questions needed",
                                "Load the questions for: %s." % ", ".join(missing))
            return
        self.data["started"] = datetime.now().isoformat(timespec="seconds")
        bump_rotation(self.data["topics"])
        self.sec = 0
        self.q_index = 0
        self.show_interview()

    # ==================================================================
    # 2. INTERVIEW
    # ==================================================================
    def _response(self, idx):
        return self.rs.setdefault(idx, {"score": None, "skipped": False, "notes": ""})

    def stats(self, topics=None):
        """Weighted by difficulty: an advanced answer counts for more than a basic one."""
        topics = self.data["topics"] if topics is None else topics
        num = den = 0.0
        raw = []
        attempted = skipped = total = 0
        by_level = {l: [] for l in LEVELS}
        for t in topics:
            qs = self.questions_for(t)
            rs = self.data["responses"].get(t, {})
            total += len(qs)
            for i, q in enumerate(qs):
                r = rs.get(i)
                if not r:
                    continue
                if r.get("skipped"):
                    score, skipped = 0, skipped + 1
                elif r.get("score") is not None:
                    score = r["score"]
                else:
                    continue
                attempted += 1
                w = LEVEL_WEIGHT.get(q["level"], 1)
                num += score * w
                den += w
                raw.append(score)
                by_level[q["level"]].append(score)
        weighted = (num / den) if den else 0.0
        return {
            "attempted": attempted,
            "total": total,
            "skipped": skipped,
            "weighted": weighted,
            "avg": (sum(raw) / len(raw)) if raw else 0.0,
            "pct": weighted / 5.0 * 100.0,
            "points": num,
            "possible": den * 5.0,
            "by_level": {l: (sum(v) / len(v)) if v else None
                         for l, v in by_level.items()},
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

        self.rail_topic = tk.Label(r, text="", bg=SURFACE, fg=ACCENT_DK,
                                   font=font(9, "bold"), anchor="w")
        self.rail_topic.pack(anchor="w", pady=(0, 10))

        tk.Label(r, text="LIVE SCORE  \u00b7  WEIGHTED", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w")
        row = tk.Frame(r, bg=SURFACE)
        row.pack(anchor="w", pady=(2, 0))
        self.rail_avg = tk.Label(row, text="0.0", bg=SURFACE, fg=ACCENT,
                                 font=font(32, "bold", "display"))
        self.rail_avg.pack(side="left")
        tk.Label(row, text="/ 5", bg=SURFACE, fg=FAINT,
                 font=font(13, "bold")).pack(side="left", anchor="s",
                                             pady=(0, 9), padx=(5, 0))
        self.rail_sub = tk.Label(r, text="", bg=SURFACE, fg=MUTED, font=font(9))
        self.rail_sub.pack(anchor="w")
        self.rail_all = tk.Label(r, text="", bg=SURFACE, fg=FAINT, font=font(9))
        self.rail_all.pack(anchor="w", pady=(2, 12))

        tk.Frame(r, bg=RULE, height=1).pack(fill="x", pady=(0, 12))
        tk.Label(r, text="PROGRESS", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 8))
        self.strip = tk.Canvas(r, bg=SURFACE, highlightthickness=0, bd=0,
                               height=130, width=228)
        self.strip.pack(anchor="w")
        self.strip.bind("<Button-1>", self._strip_click)

        tk.Frame(r, bg=RULE, height=1).pack(fill="x", pady=12)
        tk.Label(r, text="BY DIFFICULTY  \u00b7  WEIGHT", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 8))
        self.level_labels = {}
        for lvl in LEVELS:
            line = tk.Frame(r, bg=SURFACE)
            line.pack(fill="x", pady=2)
            tk.Label(line, text="%s  \u00d7%d" % (lvl, LEVEL_WEIGHT[lvl]), bg=SURFACE,
                     fg=INK, font=font(9)).pack(side="left")
            v = tk.Label(line, text="\u2013", bg=SURFACE, fg=MUTED,
                         font=font(9, "bold", "mono"))
            v.pack(side="right")
            self.level_labels[lvl] = v

        tk.Frame(r, bg=SURFACE).pack(fill="both", expand=True)
        tk.Label(r, text="A skipped question counts as 0. Advanced answers pull the "
                         "score further than basic ones.",
                 bg=SURFACE, fg=FAINT, font=font(8), wraplength=200,
                 justify="left").pack(anchor="w")

    CELL_W, CELL_H, GAP = 70, 32, 6

    def _strip_click(self, event):
        col = int(event.x // (self.CELL_W + self.GAP))
        row = int(event.y // (self.CELL_H + self.GAP))
        idx = row * 3 + col
        if 0 <= col < 3 and 0 <= idx < len(self.qs):
            self._save_current()
            self.q_index = idx
            self._render_question()

    def _draw_strip(self):
        self.strip.delete("all")
        for i, q in enumerate(self.qs):
            col, row = i % 3, i // 3
            x = col * (self.CELL_W + self.GAP)
            y = row * (self.CELL_H + self.GAP)
            r = self.rs.get(i)
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
        s = self.stats([self.topic])
        allst = self.stats()
        self.rail_topic.configure(
            text="TOPIC %d OF %d  \u00b7  %s" % (self.sec + 1,
                                                 len(self.data["topics"]), self.topic))
        self.rail_avg.configure(text="%.1f" % s["weighted"])
        self.rail_sub.configure(text="%d of %d answered  \u00b7  %d skipped"
                                     % (s["attempted"], s["total"], s["skipped"]))
        if len(self.data["topics"]) > 1:
            self.rail_all.configure(text="All topics so far: %.1f / 5" % allst["weighted"])
        for lvl, v in self.level_labels.items():
            val = s["by_level"][lvl]
            v.configure(text="\u2013" if val is None else "%.1f" % val,
                        fg=MUTED if val is None else ACCENT)
        self._draw_strip()

    # -- question ----------------------------------------------------------
    def _render_question(self):
        for w in self.stage.winfo_children():
            w.destroy()
        q = self.qs[self.q_index]
        r = self._response(self.q_index)

        outer, c = card(self.stage, pad=24)
        outer.pack(fill="both", expand=True)

        top = tk.Frame(c, bg=SURFACE)
        top.pack(fill="x")
        tk.Label(top, text="QUESTION %02d / %02d" % (self.q_index + 1, len(self.qs)),
                 bg=SURFACE, fg=MUTED, font=font(9, "bold", "mono")).pack(side="left")
        tk.Label(top, text=q["level"].upper() + "  \u00d7%d" % LEVEL_WEIGHT[q["level"]],
                 bg="#EDF4F3", fg=ACCENT_DK, font=font(8, "bold"),
                 padx=9, pady=3).pack(side="left", padx=12)
        crumb = "  \u00b7  ".join(x for x in (self.topic, q.get("subtopic"), q["qid"]) if x)
        tk.Label(top, text=crumb, bg=SURFACE, fg=FAINT,
                 font=font(9, "normal", "mono")).pack(side="right")

        tk.Label(c, text=q["question"], bg=SURFACE, fg=INK,
                 font=font(15, "bold", "display"), wraplength=660,
                 justify="center").pack(fill="x", pady=(16, 16))

        exp = tk.Frame(c, bg="#F4F7F8")
        exp.pack(fill="x", pady=(0, 14))
        tk.Frame(exp, bg=ACCENT, height=3).pack(fill="x")
        inner = tk.Frame(exp, bg="#F4F7F8", padx=14, pady=12)
        inner.pack(fill="both", expand=True)
        tk.Label(inner, text="EXPECTED ANSWER  \u00b7  INTERVIEWER ONLY", bg="#F4F7F8",
                 fg=MUTED, font=font(8, "bold")).pack(fill="x", pady=(0, 6))
        tk.Label(inner, text=q["expected"], bg="#F4F7F8", fg=INK, font=font(10),
                 wraplength=620, justify="center").pack(fill="x")

        self.notes = TextBox(c, "What the candidate said", height=4,
                             hint="optional \u00b7 goes into the record")
        self.notes.pack(fill="x", pady=(0, 14))
        self.notes.set(r["notes"])

        tk.Label(c, text="SCORE THIS ANSWER", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 6))
        self.scorebar = ScoreBar(c, command=self._on_score)
        self.scorebar.pack(anchor="w")
        if r["score"] is not None and not r["skipped"]:
            self.scorebar.set(r["score"], fire=False)

        self.skip_note = tk.Label(c, text="", bg=SURFACE, fg=WARN, font=font(9, "bold"))
        self.skip_note.pack(anchor="w", pady=(8, 0))
        if r["skipped"]:
            self.skip_note.configure(text="Marked as skipped \u2014 counts as 0. "
                                          "Pick a score to undo.")

        tk.Frame(c, bg=SURFACE).pack(fill="both", expand=True)
        tk.Frame(c, bg=RULE, height=1).pack(fill="x", pady=(14, 14))

        nav = tk.Frame(c, bg=SURFACE)
        nav.pack(fill="x")
        if self.q_index > 0:
            Button(nav, "\u2190  Previous", self._prev, kind="quiet").pack(side="left")
        Button(nav, "Skip \u2014 didn't know", self._skip,
               kind="ghost").pack(side="left", padx=8)

        last = self.q_index == len(self.qs) - 1
        more = self.sec < len(self.data["topics"]) - 1
        if last:
            label = "Finish %s  \u2192" % self.topic if more else "Finish and review  \u2192"
            Button(nav, label, self._finish_section, kind="dark").pack(side="right")
        else:
            Button(nav, "Save and next  \u2192", self._next).pack(side="right")

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
        if self.q_index < len(self.qs) - 1:
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
        if self.q_index < len(self.qs) - 1:
            self.q_index += 1
            self._render_question()

    def _prev(self):
        self._save_current()
        if self.q_index > 0:
            self.q_index -= 1
            self._render_question()

    def _finish_section(self):
        self._save_current()
        r = self._response(self.q_index)
        if r["score"] is None and not r["skipped"]:
            messagebox.showinfo("Score needed",
                                "Score the last answer, or skip it, before finishing.")
            return
        left = [i + 1 for i in range(len(self.qs))
                if self.rs.get(i, {}).get("score") is None
                and not self.rs.get(i, {}).get("skipped")]
        if left:
            if not messagebox.askyesno(
                    "Unscored questions",
                    "These %s questions have no score yet: %s.\n\nFinish anyway?"
                    % (self.topic, ", ".join("Q%d" % n for n in left))):
                return
        if self.sec < len(self.data["topics"]) - 1:
            self.show_section_done()
        else:
            self._end_test()

    def _end_test(self):
        self.data["finished"] = datetime.now().isoformat(timespec="seconds")
        self.show_wrapup()

    # -- between topics ----------------------------------------------------
    def show_section_done(self):
        self._clear()
        self._set_step(1)
        page = tk.Frame(self.body, bg=CANVAS, padx=34, pady=30)
        page.pack(fill="both", expand=True)
        done, nxt = self.topic, self.data["topics"][self.sec + 1]
        s = self.stats([done])
        allst = self.stats()

        o, c = card(page, pad=28)
        o.pack(fill="x")
        tk.Label(c, text="%s finished" % done, bg=SURFACE, fg=ACCENT,
                 font=font(22, "bold", "display")).pack(anchor="w")
        tk.Label(c, text="Topic %d of %d complete. Next up: %s."
                         % (self.sec + 1, len(self.data["topics"]), nxt),
                 bg=SURFACE, fg=MUTED, font=font(10)).pack(anchor="w", pady=(5, 20))

        m = tk.Frame(c, bg=SURFACE)
        m.pack(fill="x", pady=(0, 20))
        for label, value in ((done + " score", "%.1f / 5" % s["weighted"]),
                             ("Answered", "%d of %d" % (s["attempted"], s["total"])),
                             ("Skipped", str(s["skipped"])),
                             ("All topics so far", "%.1f / 5" % allst["weighted"])):
            bo, b = card(m, pad=15, bg="#F5F8F9", border="#E6ECEE")
            bo.pack(side="left", fill="x", expand=True, padx=(0, 12))
            tk.Label(b, text=label.upper(), bg="#F5F8F9", fg=MUTED,
                     font=font(8, "bold")).pack(anchor="w")
            tk.Label(b, text=value, bg="#F5F8F9", fg=INK,
                     font=font(18, "bold", "display")).pack(anchor="w", pady=(2, 0))

        for lvl in LEVELS:
            val = s["by_level"][lvl]
            line = tk.Frame(c, bg=SURFACE)
            line.pack(fill="x", pady=1)
            tk.Label(line, text="%s average" % lvl, bg=SURFACE, fg=MUTED,
                     font=font(9)).pack(side="left")
            tk.Label(line, text="\u2013" if val is None else "%.1f / 5" % val,
                     bg=SURFACE, fg=INK, font=font(9, "bold")).pack(side="left", padx=8)

        tk.Frame(c, bg=RULE, height=1).pack(fill="x", pady=(20, 18))
        row = tk.Frame(c, bg=SURFACE)
        row.pack(fill="x")
        Button(row, "\u2190  Back to %s" % done, self._back_to_section,
               kind="quiet").pack(side="left")
        Button(row, "Exit the test", self._exit_test, kind="ghost").pack(side="left", padx=8)
        Button(row, "Next topic: %s  \u2192" % nxt, self._next_section,
               kind="dark").pack(side="right")

    def _back_to_section(self):
        self.q_index = max(0, len(self.qs) - 1)
        self.show_interview()

    def _next_section(self):
        self.sec += 1
        self.q_index = 0
        self.show_interview()

    def _exit_test(self):
        remaining = self.data["topics"][self.sec + 1:]
        if not messagebox.askyesno(
                "Exit the test",
                "End the interview here?\n\n%s will not be asked. You still go on "
                "to the rating and the documents."
                % ", ".join(remaining)):
            return
        self.data["ended_early"] = True
        self._end_test()

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
            cand.get("candidate", ""), " + ".join(self.data["topics"]),
            self.data["interview_id"]),
                 bg=CANVAS, fg=MUTED, font=font(10)).pack(anchor="w", pady=(4, 20))

        metrics = tk.Frame(page, bg=CANVAS)
        metrics.pack(fill="x", pady=(0, 18))
        for label, value in (("Weighted score", "%.1f / 5" % s["weighted"]),
                             ("Percentage", "%d%%" % round(s["pct"])),
                             ("Answered", "%d of %d" % (s["attempted"], s["total"])),
                             ("Skipped", str(s["skipped"]))):
            o, b = card(metrics, pad=16)
            o.pack(side="left", fill="x", expand=True, padx=(0, 12))
            tk.Label(b, text=label.upper(), bg=SURFACE, fg=MUTED,
                     font=font(8, "bold")).pack(anchor="w")
            tk.Label(b, text=value, bg=SURFACE, fg=INK,
                     font=font(20, "bold", "display")).pack(anchor="w", pady=(2, 0))

        # --- per topic ---
        o0, tc = card(page)
        o0.pack(fill="x", pady=(0, 18))
        tk.Label(tc, text="By topic", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w", pady=(0, 12))
        for t in self.data["topics"]:
            ts = self.stats([t])
            line = tk.Frame(tc, bg=SURFACE)
            line.pack(fill="x", pady=3)
            tk.Label(line, text=t, bg=SURFACE, fg=INK, font=font(10, "bold"),
                     width=12, anchor="w").pack(side="left")
            if ts["attempted"] == 0:
                tk.Label(line, text="Not asked", bg=SURFACE, fg=FAINT,
                         font=font(9)).pack(side="left")
            else:
                tk.Label(line, text="%.1f / 5" % ts["weighted"], bg=SURFACE, fg=ACCENT,
                         font=font(10, "bold")).pack(side="left")
                tk.Label(line, text="   %d answered  \u00b7  %d skipped  \u00b7  %s"
                                    % (ts["attempted"], ts["skipped"],
                                       "  ".join("%s %s" % (l[:3],
                                                            "\u2013" if ts["by_level"][l] is None
                                                            else "%.1f" % ts["by_level"][l])
                                                 for l in LEVELS)),
                         bg=SURFACE, fg=MUTED, font=font(9)).pack(side="left")

        # --- interviewer assessment ---
        o, c = card(page)
        o.pack(fill="x", pady=(0, 18))
        tk.Label(c, text="Your assessment", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w")
        tk.Label(c, text="Your own rating, independent of the calculated score.",
                 bg=SURFACE, fg=MUTED, font=font(9)).pack(anchor="w", pady=(3, 12))

        tk.Label(c, text="OVERALL INTERVIEWER RATING", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 6))
        self.rating_bar = ScoreBar(c)
        self.rating_bar.pack(anchor="w", pady=(0, 16))
        if self.data["rating"]:
            self.rating_bar.set(self.data["rating"], fire=False)

        tk.Label(c, text="RECOMMENDATION", bg=SURFACE, fg=MUTED,
                 font=font(8, "bold")).pack(anchor="w", pady=(0, 8))
        pills = tk.Frame(c, bg=SURFACE)
        pills.pack(anchor="w", pady=(0, 16))
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

        self.comment_box = TextBox(c, "Comments", height=5, hint="goes into the report")
        self.comment_box.pack(fill="x")
        self.comment_box.set(self.data["comment"])

        # --- review ---
        o2, rv = card(page)
        o2.pack(fill="x", pady=(0, 18))
        tk.Label(rv, text="Answer review", bg=SURFACE, fg=INK,
                 font=font(12, "bold", "display")).pack(anchor="w", pady=(0, 10))
        for t in self.data["topics"]:
            qs = self.questions_for(t)
            if not qs:
                continue
            tk.Label(rv, text=t.upper(), bg=SURFACE, fg=ACCENT_DK,
                     font=font(9, "bold")).pack(anchor="w", pady=(10, 4))
            rs = self.data["responses"].get(t, {})
            for i, q in enumerate(qs):
                r = rs.get(i, {})
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
                         anchor="w", wraplength=740, justify="left").pack(side="left",
                                                                          fill="x")
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
        self.q_index = 0
        self.show_interview()

    def _goto_export(self):
        self._harvest_wrapup()
        if not self.data["rating"]:
            messagebox.showinfo("Rating needed",
                                "Give an overall rating from 1 to 5 before continuing.")
            return
        if not self.data["decision"]:
            messagebox.showinfo("Recommendation needed",
                                "Pick a recommendation before continuing.")
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
        tk.Label(page, text="Two files are written: a plain transcript for the agent, "
                            "and a formatted PDF report.",
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
        tk.Label(c, text="A folder named %s is created here." % self.data["interview_id"],
                 bg=SURFACE, fg=FAINT, font=font(9)).pack(anchor="w", pady=(8, 0))

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
        if hasattr(self, "ans_var"):
            self.data["include_answers"] = bool(self.ans_var.get())
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
            p1 = os.path.join(folder, prefix + "_Transcript.docx")
            self._build_interview_transcript().save(p1)
            made.append(p1)

            p2 = os.path.join(folder, prefix + "_Report.pdf")
            self._build_report().save(p2)
            made.append(p2)

            src = self.data.get("transcript_source")
            if src and os.path.isfile(src):
                p3 = os.path.join(folder, prefix + "_MeetingTranscript.docx")
                self._build_meeting_transcript(src).save(p3)
                made.append(p3)
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
        tk.Label(c, text="Next candidate keeps your team, role, grade, requisition ID "
                         "and name, the same topics in the same order, and the same "
                         "question files \u2014 all of it still editable. Only the "
                         "candidate's own details are cleared, and every topic moves "
                         "on to the next set of questions.",
                 bg=SURFACE, fg=FAINT, font=font(9), wraplength=820,
                 justify="left").pack(anchor="w", pady=(22, 10))

        row = tk.Frame(c, bg=SURFACE)
        row.pack(anchor="w", pady=(2, 0))
        Button(row, "Open folder", lambda: open_folder(folder), kind="primary").pack(side="left")
        Button(row, "Next candidate", self._new_interview, kind="dark").pack(side="left", padx=10)
        Button(row, "Exit", self._exit, kind="ghost").pack(side="left")

    def _new_interview(self):
        if not messagebox.askyesno(
                "Next candidate",
                "Clear this interview and set up the next one?\n\n"
                "The documents already saved are not affected."):
            return
        old = dict(self.data.get("candidate") or {})
        now = datetime.now()
        keep = {k: old.get(k, "") for k in ("team", "role", "grade", "req", "interviewer")}
        keep["date"] = now.strftime("%d %b %Y")
        keep["time"] = now.strftime("%I:%M %p").lstrip("0")
        topics = list(self.data.get("topics") or [])
        banks = {t: dict(b) for t, b in (self.data.get("banks") or {}).items()}
        self.reset_data()
        self.data["candidate"] = keep
        self.data["topics"] = topics
        self._set_id_chip()
        self.show_setup()
        for t in topics:
            if t in banks:
                self._rebuild_bank(t, banks[t])
        self._paint_tiles()
        self._render_topic_rows()

    def _exit(self):
        if messagebox.askyesno("Close Interview Assistant",
                               "Close the app?\n\nAnything not yet generated is lost."):
            self.destroy()

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

    def _mark_text(self, r):
        if r.get("skipped"):
            return "0 out of 5"
        if r.get("score") is None:
            return "Not scored"
        return "%d out of 5" % r["score"]

    def _answer_text(self, r):
        if r.get("skipped"):
            return "Skipped - the candidate did not know the answer."
        return r.get("notes") or "Not recorded."

    def _build_interview_transcript(self):
        """Plain text: the questions, the answers, and the mark for each."""
        d = Docx()
        cand = self.data["candidate"]
        answers = self.data.get("include_answers", True)

        def L(text="", bold=False):
            d.para(text, bold=bold, size=11, space_after=0)

        L("INTERVIEW TRANSCRIPT", bold=True)
        L()
        L("Interview ID: %s" % self.data["interview_id"])
        L("Candidate: %s" % (cand.get("candidate") or "Not recorded"))
        L("Date: %s" % (cand.get("date") or "Not recorded"))
        L("Topics: %s" % ", ".join(self.data["topics"]))
        L()

        for t in self.data["topics"]:
            qs = self.questions_for(t)
            rs = self.data["responses"].get(t, {})
            if not qs:
                continue
            L("TOPIC: %s" % t.upper(), bold=True)
            L()
            for i, q in enumerate(qs):
                r = rs.get(i, {})
                L("Q%d (%s)" % (i + 1, q["level"]), bold=True)
                if q.get("subtopic"):
                    L("Sub-topic: %s" % q["subtopic"])
                L("Question: %s" % q["question"])
                if answers:
                    L("Candidate answer: %s" % self._answer_text(r))
                L("Mark: %s" % self._mark_text(r))
                L()
        L("End of transcript.")
        return d

    # -- PDF report --------------------------------------------------------
    def _pdf_band(self, d):
        h = 92
        d.rect(0, d.H - h, d.W, h, HEADER_RGB)
        d.rect(0, d.H - h, d.W, 3, ACCENT_RGB)
        d.text(d.margin, d.H - 44, "Interview Assessment Report", 19, True, (255, 255, 255))
        d.text(d.margin, d.H - 64,
               "%s  |  %s" % (self.data["candidate"].get("candidate") or "Candidate",
                              self.data["interview_id"]), 9.5, False, (168, 200, 205))
        d.y = d.H - h - 26

    def _pdf_stats(self, d, items):
        h, gap = 54, 10
        d.need(h + 10)
        w = (d.usable - gap * (len(items) - 1)) / float(len(items))
        top = d.y
        for i, (label, value) in enumerate(items):
            x = d.margin + i * (w + gap)
            d.rect(x, top - h, w, h, PANEL_RGB)
            d.rect(x, top - h, 2.5, h, ACCENT_RGB)
            d.text(x + 12, top - 19, label.upper(), 7, True, MUTED_RGB)
            size = 15 if d.width(value, 15, True) < w - 24 else 10.5
            d.text(x + 12, top - 40, value, size, True, INK_RGB)
        d.y = top - h - 4

    def _pdf_kv(self, d, pairs, cols=3):
        rows = (len(pairs) + cols - 1) // cols
        w = d.usable / float(cols)
        top = d.y
        for i, (k, v) in enumerate(pairs):
            col, row = i % cols, i // cols
            x = d.margin + col * w
            y = top - row * 30
            d.text(x, y - 9, k.upper(), 7, True, MUTED_RGB)
            d.text(x, y - 21, v or "Not recorded", 9.5, False, INK_RGB)
        d.y = top - rows * 30

    def _pdf_topic_table(self, d):
        cols = [0.30, 0.16, 0.16, 0.18, 0.20]
        head = ["Topic", "Asked", "Skipped", "Score / 5", "Percentage"]
        xs, run = [], d.margin
        for c in cols:
            xs.append(run)
            run += c * d.usable
        d.need(26)
        d.rect(d.margin, d.y - 20, d.usable, 20, PANEL_RGB)
        for x, htext in zip(xs, head):
            d.text(x + 7, d.y - 14, htext, 7.5, True, MUTED_RGB)
        d.y -= 20
        for t in self.data["topics"]:
            s = self.stats([t])
            asked = s["attempted"]
            vals = [t,
                    "%d of %d" % (asked, s["total"]),
                    str(s["skipped"]),
                    "-" if not asked else "%.1f" % s["weighted"],
                    "-" if not asked else "%d%%" % round(s["pct"])]
            d.need(20)
            for x, v in zip(xs, vals):
                d.text(x + 7, d.y - 14, v, 9, False, INK_RGB)
            d.y -= 20
            d.line(d.margin, d.y, d.right, d.y, RULE_RGB, 0.5)
        d.y -= 4

    def _build_report(self):
        cand = self.data["candidate"]
        s = self.stats()
        answers = self.data.get("include_answers", True)
        d = Pdf(footer="Confidential  |  %s  |  %s"
                       % (self.data["interview_id"],
                          cand.get("candidate") or "Candidate"))
        self._pdf_band(d)

        d.para("This report records a structured technical interview. Questions were "
               "drawn from a fixed bank across three difficulty bands, scored from 1 "
               "to 5 as the interview ran. The overall score is weighted by difficulty, "
               "so advanced answers count for more than basic ones.",
               9.5, color=MUTED_RGB, after=16)

        self._pdf_stats(d, [
            ("Weighted score", "%.1f / 5" % s["weighted"]),
            ("Percentage", "%d%%" % round(s["pct"])),
            ("Interviewer rating", "%s / 5" % (self.data["rating"] or "-")),
            ("Recommendation", self.data["decision"] or "Not recorded"),
        ])

        d.heading("Candidate and interview details", before=18)
        self._pdf_kv(d, [
            ("Candidate", cand.get("candidate")),
            ("Candidate ID", cand.get("cid")),
            ("Interview ID", self.data["interview_id"]),
            ("Role", cand.get("role")),
            ("Grade", cand.get("grade")),
            ("Requisition", cand.get("req")),
            ("Team", cand.get("team")),
            ("Date and time", "%s  %s" % (cand.get("date", ""), cand.get("time", ""))),
            ("Interviewer", cand.get("interviewer")),
            ("Topics", ", ".join(self.data["topics"])),
            ("Questions asked", "%d of %d" % (s["attempted"], s["total"])),
            ("Duration", self._duration()),
        ])

        d.heading("Performance by topic")
        self._pdf_topic_table(d)

        d.heading("Performance by difficulty")
        band = []
        for lvl in LEVELS:
            v = s["by_level"][lvl]
            band.append(("%s (weight %d)" % (lvl, LEVEL_WEIGHT[lvl]),
                         "-" if v is None else "%.1f / 5" % v))
        self._pdf_stats(d, band)

        d.heading("Interviewer assessment")
        self._pdf_stats(d, [
            ("Interviewer rating", "%s out of 5" % (self.data["rating"] or "-")),
            ("Recommendation", self.data["decision"] or "Not recorded"),
        ])
        d.space(8)
        d.para("COMMENT", 7, bold=True, color=MUTED_RGB, after=1)
        d.para(self.data["comment"] or "No comment was recorded.", 9.5,
               color=MUTED_RGB, after=4)

        d.heading("Question detail")
        for t in self.data["topics"]:
            qs = self.questions_for(t)
            rs = self.data["responses"].get(t, {})
            if not qs:
                continue
            d.need(40)
            d.y -= 6
            d.text(d.margin, d.y - 10, t.upper(), 9, True, ACCENT_RGB)
            d.y -= 22
            for i, q in enumerate(qs):
                r = rs.get(i, {})
                block = 32 + len(d.wrap(q["question"], 9.5, d.usable - 74)) * 13
                d.need(block)
                d.text(d.margin, d.y - 10, "Q%d" % (i + 1), 8.5, True, MUTED_RGB)
                d.text(d.margin + 26, d.y - 10, q["level"], 7.5, False, FAINT_RGB)
                mark = self._mark_text(r)
                d.text(d.right - d.width(mark, 8.5, True), d.y - 10, mark, 8.5, True,
                       INK_RGB)
                d.y -= 14
                d.para(q["question"], 9.5, color=INK_RGB, indent=26, lead=1.35)
                if answers:
                    d.para("Answer: %s" % self._answer_text(r), 8.5,
                           color=MUTED_RGB, indent=26, lead=1.35)
                d.y -= 5
                d.line(d.margin, d.y, d.right, d.y, RULE_RGB, 0.5)
                d.y -= 5

        if self.data.get("ended_early"):
            d.para("Note: the interview was ended before every selected topic was "
                   "asked. Topics with no questions asked are shown as not asked.",
                   8.5, color=FAINT_RGB, after=4)
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
        L("Topics: %s" % ", ".join(self.data["topics"]))
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
