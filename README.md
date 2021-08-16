# ExcelToLaTeX
Convert Excel-Sheets to LaTeX tables via a Python interface.

# Usage
Load .xlsx-file with 
```python
dataframe = load_excel('file.xlsx', column_names=['column1', 'column2'])
```
Afterwards, you can convert the dataframe to LaTeX by using:
```python
parse(dataframe, 'table.txt')
```
The output will be saved into the specified file:
```latex
\begin{table}[H]
\centering
\begin{tabular}{ll}
\toprule
column1 & column2\\ 
\midrule
\rowcolor{gray} $4.0$ & $1.0$\\
$4.5$ & $1.2$\\
\rowcolor{gray} $4.8$ & $1.3$\\
$5.2$ & $1.5$\\
\rowcolor{gray} $6.0$ & $1.8$\\
\bottomrule
\end{tabular}
\caption{}
\end{table}
```
