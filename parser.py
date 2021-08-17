import os
import pandas as pd 


def load_excel(filename: str, path : str = os.getcwd(), **kwargs) -> pd.DataFrame:

	"""
	Load the .xlsx-file via pandas into pd.DataFrame format.

	args:
		filename (str): name of .xlsx-file
		path (str, optional): absolute path to file
		**kwargs: see pandas.read_excel documentation for additional arguments

	returns:
		pd.DataFrame: dataframe with data
	"""

	if not os.path.isfile(os.path.join(path, filename)):
		raise OSError(f"File '{os.path.join(path, filename)}' not found.")

	dataframe = pd.read_excel(os.path.join(path, filename), **kwargs)
	
	return dataframe


def generate_header(orientation: list, number_columns: int, complete_document: bool, disable_debug: bool) -> str:

	"""
	Generate table header. 

	args:
		orientation (list): orientation of individual columns (left | center | right). 
							If only one format is specified it will be applied to all columns.
		number_columns (int): number of columns

	returns:
		str: table header
	"""

	ORIENTATION = {'left': 'l', 'center': 'c', 'right': 'r'}
	column_orientation = ''

	if number_columns != len(orientation):

		for index in range(number_columns):
			column_orientation += ORIENTATION[orientation[0]]

	else:

		for index, column in enumerate(orientation):
			column_orientation += ORIENTATION[column.lower()]

	if not disable_debug and not complete_document:
		return '''% Include these packages\n% Figure Orientation\n% \\usepackage{float}\n% Booktabs for nice tables\n% \\usepackage{booktabs}\n% color for row coloring\n% \\usepackage{xcolor, colortbl}\n% \\definecolor{gray}{rgb}{0.85, 0.85, 0.85}\n\n\\begin{table}[H]\n\\centering\n\\begin{tabular}{''' + column_orientation + '''}\n'''

	elif complete_document:
		return '''\\documentclass[a4paper, 12pt]{article}\n\n\\usepackage[utf8]{inputenc}\n\\usepackage[T1]{fontenc}\n\\usepackage[english]{babel}\n\\usepackage[a4paper, left=2.5cm, right=2.5cm, top=2.5cm, bottom=3cm]{geometry}\n\\usepackage{float}\n\\usepackage{booktabs}\n\\usepackage{xcolor, colortbl}\n\\definecolor{gray}{rgb}{0.85, 0.85, 0.85}\n\n\\begin{document}\n\n\\begin{table}[H]\n\\centering\n\\begin{tabular}{''' + column_orientation + '''}\n''' 

	else:
		return '''\\begin{table}[H]\n\\centering\n\\begin{tabular}{''' + column_orientation + '''}\n'''


def generate_body(dataframe: pd.DataFrame, striped: bool = True, is_numeric: bool = True, decimal_sep: str = '.') -> str:

	"""
	Generate table body.

	args:
		dataframe (pd.DataFrame): 
		striped (bool, optional): True for striped row color
		is_numeric (bool, optional): columns contain numeric values, additional math mode signs will be added
		decimal_sep (str, optional): convert decimal separator (decimal point as default)

	returns:
		str: table body
	"""

	body = '''\\toprule\n'''

	column_names = dataframe.columns

	for index, column in enumerate(column_names):
		body += column 

		if index == (len(column_names) - 1):
			body += '\\\\ \n\\midrule\n'
		else:
			body += ' & '

	for row_index, row in dataframe.iterrows():

		if (row_index % 2) == 0:
			body += '\\rowcolor{gray} '

		for index, item in enumerate(row):

			if is_numeric:
				body += "$%s$" % str(item).replace('.', decimal_sep)
			else:
				body += str(item).replace('.', decimal_sep)

			if index == (len(row) - 1):
				body += '\\\\\n'

			else:
				body += ' & '

	return body + '\\bottomrule\n'


def generate_footer(caption: str = '', complete_document: bool = False) -> str:

	"""
	Generate table footer.

	args:
		caption (str, optional): table caption, blank if not specified

	returns:
		str: table footer
	"""

	if not complete_document:
		return '''\\end{tabular}\n\\caption{''' + caption + '''}\n\\end{table}'''

	else:
		return '''\\end{tabular}\n\\caption{''' + caption + '''}\n\\end{table}\n\n\\end{document}'''


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

	table = generate_header(orientation, len(dataframe.columns), complete_document=complete_document, disable_debug=disable_debug) + \
			generate_body(dataframe, striped=striped, is_numeric=is_numeric, decimal_sep=decimal_sep) + \
			generate_footer(caption=caption, complete_document=complete_document)

	if os.path.isfile(os.path.join(path, filename)) and not overwrite:
		raise IOError(f'File {os.path.join(path, filename)} already exists. Specify a different filename or set overwrite=True.')

	try:

		with open(os.path.join(path, filename), 'w') as file:
			file.write(table)

	except Exception as exception:
		raise exception


if __name__ == '__main__':
	
	pass

