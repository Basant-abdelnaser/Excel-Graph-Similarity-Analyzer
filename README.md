# File Similarity Analyzer

A C# application for analyzing and clustering files based on similarity using **graph algorithms**.  
The tool parses Excel datasets, constructs a graph representation, detects connected components, and applies **Maximum Spanning Tree (MST)** to refine relationships.  
Results are exported as detailed **CSV reports** containing similarity statistics and refined file-pair matches.

---

## Features
-  **Excel Parsing:** Extracts file-pair data using **NPOI** (supports `.xlsx` and `.xls` formats).  
-  **Graph Construction:** Represents file similarity as an undirected graph.  
-  **Connectivity Detection:** Uses **BFS** to detect connected components.  
-  **Redundancy Removal:** Builds a **Maximum Spanning Tree (MST)** for each component.  
-  **Reporting:** Generates CSV reports with:
  - Connected groups of files  
  - Average similarity per component  
  - Refined file-pair matches  
-  **Performance Profiling:** Measures execution times of different stages.

---

##  Technologies
- **Language:** C# (.NET)  
- **Libraries:** [NPOI](https://github.com/nissl-lab/npoi) for Excel handling  
- **Algorithms:** BFS, MST (Kruskal-based approach)

---

##  Input
An Excel sheet with file-pair data. Example format:

| File 1 (path + similarity) | File 2 (path + similarity) | Line Matches |
|-----------------------------|-----------------------------|--------------|
| fileA(85)                   | fileB(78)                   | 120          |
| fileB(78)                   | fileC(80)                   | 95           |

---

## Output
Two CSV reports are generated:
1. **export-stat.csv** – Group statistics with average similarity and component count  
2. **export.csv** – Refined file-pair matches (MST output)  

---

## How to Run
1. Clone the repository:
   ```bash
   git clone https://github.com/<your-username>/<repo-name>.git
