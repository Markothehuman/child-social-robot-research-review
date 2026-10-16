CRI Preliminary Corpus Audit — Updated Zotero Export

This folder contains the current working audit of the updated Zotero bibliography.

FILES
00_Source_Zotero_Export.bib
    The updated Zotero export used for this audit.

01_CRI_Preliminary_Corpus_Inventory.csv
    Running paper list with automated first-pass evidence type, setting,
    methodology, topic and population tags plus empty manual screening fields.

02_CRI_Preliminary_Corpus_Summary.md
    Short corpus-level summary suitable for preparing supervisor updates.

03_CRI_Potential_Duplicates.csv
    Exact DOI/title duplicate groups flagged for manual cleanup.

04_cri_corpus_mapper.py
    Lightweight reusable helper / record of the mapping workflow.

05_CRI_Preliminary_Corpus_Inventory.xlsx
    Editable Excel version with Summary and Inventory sheets.

06_cri_corpus_summary.py
    Summarises the existing Inventory tags and generates five simple charts.
    Does not classify papers, change tags, or remove potential duplicates.

06_cri_corpus_summary.ipynb
    Executed Jupyter notebook version with the charts and summary embedded.
    Open in VS Code or Jupyter, choose the Anaconda Python kernel, then Run All
    to refresh it. Counting and plotting helpers are included in the notebook.
    Start the kernel in this folder, or adjust the workbook path in Setup.
    Also saves the same five PNG charts and corpus_summary.txt.

requirements.txt
    Python dependencies for the summary script and notebook kernel.

RUN THE CORPUS SUMMARY
Install dependencies if needed:
    python -m pip install -r requirements.txt

Save the charts and also display them in plot windows:
    python 06_cri_corpus_summary.py

Save without opening plot windows:
    python 06_cri_corpus_summary.py --no-show

The outputs in this folder were generated using the installed Anaconda Python.
If the python command does not work in PowerShell on this machine, use:
    & "$env:USERPROFILE\anaconda3\python.exe" .\06_cri_corpus_summary.py --no-show
Omit --no-show to display the figures as well.

If a notebook kernel reports that openpyxl is missing, run this in a notebook
cell to install it into that selected kernel:
    %pip install openpyxl
Then choose Restart Kernel and Run All. The setup cell prints the actual
kernel interpreter and Excel-reader version, and refreshes import caches.

Optional: supply another workbook, sheet, or output directory:
    python 06_cri_corpus_summary.py "updated_inventory.xlsx" --sheet Inventory --output-dir "corpus_summary_outputs" --no-show

By default, outputs are saved in corpus_summary_outputs beside the workbook:
    evidence_type_breakdown.png
    settings_breakdown.png
    methodologies_breakdown.png
    topics_breakdown.png
    publications_by_year.png
    corpus_summary.txt

The script prints the detected sheet and column names. Category names come
directly from the spreadsheet. Semicolon-separated tags are counted once per
record per category; one record may belong to several categories. Charts and
text show untagged record counts. Duplicate groups use existing dedupe_key
flags (or the duplicates CSV if that column is absent), without deleting rows.
Publication years are preserved as recorded, including future-dated records.
Journal/conference counts use the existing bibliographic record type.
These are descriptive corpus counts, not final narrative-review findings.

CURRENT SNAPSHOT
- 116 Zotero records
- +10 records compared with the previous export
- 1 exact DOI/title duplicate group(s) flagged

CAUTION
All thematic/methodological classifications are preliminary automated tags
from bibliographic metadata. Verify them manually from the full papers,
especially study setting, methodology and evidence type.
