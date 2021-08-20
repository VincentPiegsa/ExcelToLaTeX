# ExcelToLaTeX
Convert Excel-Sheets to LaTeX tables via a Python interface.

# Get Started
## Necessary modules
Firstly, [Python](https://www.python.org/downloads/) has to be installed if it is not already. Before running the script, [pandas](https://pandas.pydata.org/) has to be installed via [pip](https://pypi.org/project/pip/):
```bash
python -m pip install pandas --upgrade
```

## Generate a table
Load the Excel-file containing the data with:
```python
dataframe = load_excel('file.xlsx')
```
The first row of the table will be used as column titles. Afterwards, you can convert the dataframe to LaTeX by using:
```python
parse(dataframe, 'table.tex', is_numeric=True, caption='Nicely formatted table')
```
The output will be saved into the specified file, that is located in the same directory where the script is executed or in the specified path:
```latex
\begin{table}[H]
	\centering
	\begin{tabular}{lll}
		\toprule
		Voltage & Current  & Resistance\\ 
		\midrule
		\rowcolor{gray} $3,1$ & $0,5$ & $6,2$\\
		$5,9$ & $1,0$ & $5,9$\\
		\rowcolor{gray} $8,8$ & $1,5$ & $5,87$\\
		$12,3$ & $2,0$ & $6,15$\\
		\rowcolor{gray} $15,0$ & $2,5$ & $6,0$\\
		$17,8$ & $3,0$ & $5,93$\\
		\bottomrule
	\end{tabular}
	\caption{Nicely formatted table}
\end{table}
```
Finally, the generated table code can be included into the main .tex-document with:
```latex
\input{table.tex}
```
![Image](https://github.com/VincentPiegsa/ExcelToLaTeX/blob/main/docs/table.pdf)
Note: to use the full capabilites of the table parser, some additional libraries have to be included in the preamble:
```latex
% figure orientation
\usepackage{float}
% booktabs for nice tables
\usepackage{booktabs}
% color for striped rows
\usepackage{xcolor, colortbl}
\definecolor{gray}{rgb}{0.85, 0.85, 0.85}
```

# Documentation
1. Load Excel-file: input the .xlsx-file. If the file is located in a different directory, also specify the absolute path. For additional arguments, look up the [pandas.read_excel documentation](https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html).
```python
load_excel(filename: str, path : str = os.getcwd(), **kwargs) -> pd.DataFrame:
	"""
	Load the .xlsx-file via pandas into pd.DataFrame format.

	args:
		filename (str): name of .xlsx-file
		path (str, optional): absolute path to file, current directory by default
		**kwargs: see pandas.read_excel documentation for additional arguments: 

	returns:
		pd.DataFrame: dataframe with data
	"""
```
2. Parse dataframe to LaTeX-format: input the pandas dataframe, and specify the filename of the output file. If neccesary, also add the absolute path, otherwise the current directory is parsed. Further optional arguments:
* orientation: orientation of the individual columns. Use "right", "center" or "left", or just specify one orientation to be applied for all columns.
* caption: table caption, empty string if not specified.
* striped: striped rows (every second row is colored gray by default).
* is_numeric: True if column contains numeric values. For the final output, $-signs are added to enter mathmode.
* decimal_sep: decimal separator, "," if not specified.
* overwrite: overwrite output file if already exists.
* complete_document: if True add a minimalist LaTeX document around the table
* disable_debug: if True suggested packages won't be shown in the output file
```python
def parse(dataframe: pd.DataFrame, filename: str, path: str = os.getcwd(), orientation: list = ['left'], caption: str = 'Table', striped: bool = True, is_numeric: bool = False, decimal_sep: str = ',', overwrite: bool = False, complete_document: bool = False, disable_debug: bool = False) -> None:

	"""
	Parse pandas dataframe to LaTeX table format.
	args:
		dataframe (pd.DataFrame): dataframe with data
		filename (str): filename of output file
		path (str, optional): path to output file
		orientation (list, optional): orientation of individial columns
		caption (str, optoinal): table caption
		striped (bool, optional): True if rows with striped color
		is_numeric (bool, optional): True if columns contain numeric values
		decimal_sep (str, optional): specify decimal separator
		overwrite (bool, optional): overwrite output file if already exists
		complete_document (bool, optional): if True outputs a minimalist LaTeX document
		disable_debug (bool, optional): if True package-info won't be written into the output file
	"""
```
