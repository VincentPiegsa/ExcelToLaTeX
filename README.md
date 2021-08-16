# ExcelToLaTeX
Convert Excel-Sheets to LaTeX tables via a Python interface.

# Get Started
Load .xlsx-file with 
```python
dataframe = load_excel('file.xlsx')
```
The first row of the table will be used as column titles. Afterwards, you can convert the dataframe to LaTeX by using:
```python
parse(dataframe, 'table.txt')
```
The output will be saved into the specified file, that is located in the same directory where the script is executed:
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

# Documentation
Load Excel-file: input the .xlsx-file. If the file is located in a different directory, also specify the absolute path. For additional arguments, look up the [pandas.read_excel documentation](https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html).
```python
load_excel(filename: str, path : str = '', **kwargs) -> pd.DataFrame:
	"""
	Load the .xlsx-file via pandas into pd.DataFrame format.

	args:
		filename (str): name of .xlsx-file
		path (str, optional): absolute path to file
		**kwargs: see pandas.read_excel documentation for additional arguments: 

	returns:
		pd.DataFrame: dataframe with data
	"""
```
Parse dataframe to LaTeX-format: input the pandas dataframe, and specify the filename of the output file. If neccesary, also add the absolute path, otherwise the current directory is parsed. 
```python
parse(dataframe: pd.DataFrame, filename: str, path: str = os.getcwd(), orientation: list = ['left'], caption: str = 'Table', striped: bool = True, is_numeric: bool = False, decimal_sep: str = ',', overwrite: bool = False) -> None:
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
		
	returns:
		None
	"""
```
