import os, shutil

def analyze_filename(filename):
	# In main folder, filename "04. Định lượng. a. Hồ sơ.doc" should be 
	# processed as (4, 'Định lượng', 'a', 'Hồ sơ.doc')
	# Number 4 refers to "Định lượng"
	# Letter "a" refers to "Hồ sơ"
	# In ./Data folder, filename "04. Định lượng. e. Thử 1.pdf" should be 
	# processed as (4, 'Định lượng', 'e', 'Thử 1.pdf')
	# If filename can't be processed as above, the file considered as 
	# not the part of document.
	
	_split = filename.split(sep='. ', maxsplit=4)
	try:
		_order = int(_split[0])
		_sub_order = _split[2]
		return (_order, _split[1], _sub_order, _split[3])
	except:
		return None


class AnalytDir:
	def __init__(self, path):
		self.path = path.replace('/', '\\') + '\\' * (not path.endswith('\\'))
		self.filelist = []
		self.valid_main = []
		self.valid_main_analyzed = []
		self.valid_extra = []
		self.valid_extra_analyzed = []
		self.valid_analyzed = []
		self.max_order = 0
		self.max_subchar = {}	# Highest letter of an order, eg {4: 'c']}
		self.spec_orders = {}	# Order of a spec, eg {'Định lượng': 4}
		self.analyze_folder()
	
	def analyze_folder(self):
		"""Analyze files in main and Main/Data folders
		If filename is formatted properly, add it to valid list
		"""
		if os.path.exists(self.path):
			_main_files = [f.path.split('\\')[-1] \
				for f in os.scandir(self.path) if f.is_file()]
			for f in _main_files:
				a = analyze_filename(f)
				if a is not None:
					self.valid_main.append(f)
					self.valid_main_analyzed.append(a)
		else:
			_main_files = []
		
		if os.path.exists(self.path + '\\Data'):
			_extra_files = [f.path.split('\\')[-1] \
				for f in os.scandir(self.path + '\\Data') if f.is_file()]
			for f in _extra_files:
				a = analyze_filename(f)
				if a is not None:
					self.valid_extra.append(f)
					self.valid_extra_analyzed.append(a)
		else:
			_extra_files = []
		
		self.filelist = _main_files + _extra_files
		self.valid_analyzed = self.valid_main_analyzed + self.valid_extra_analyzed
		for a in self.valid_analyzed:
			if a[0] in self.max_subchar.keys():
				if a[2] > self.max_subchar[a[0]]:
					self.max_subchar[a[0]] = a[2]
			else:
				self.max_subchar[a[0]] = a[2]
			
			if a[0] not in self.spec_orders.values():
				self.spec_orders[a[1]] = a[0]
			
			if a[0] > self.max_order:
				self.max_order = a[0]
	
	def add_file(self, filepath, spec_name, doc_type, del_src=False, overwrite=False):
		"""Add a file to main or data folder
		Rename file to defined format
		If doc_type = 'doc', the file considered as a document file, subprefix = 'a', file named as 'Hồ sơ'
		If doc_type = 'ss', the file considered as a spreadsheet, subprefix = 'b', file named as 'Bảng tính'
		If doc_type = 'ex', the file considered as a extra data, subprefix count on from 'c', file name keeps unchanged
		"""
		print(filepath)
		if not os.path.isfile(filepath):
			return 1
			
		if spec_name in self.spec_orders.keys():
			spec_order = self.spec_orders[spec_name]
		else:
			spec_order = self.max_order + 1

		if doc_type == 'doc':	# Subprefix = 'a'
			os.makedirs(self.path, exist_ok=True)
			_new_file = '. '.join([f'{spec_order:02}', spec_name, 'a', 'Hồ sơ' + os.path.splitext(filepath)[1]])
		elif doc_type == 'ss':	# # Subprefix = 'b'
			os.makedirs(self.path, exist_ok=True)
			_new_file = '. '.join([f'{spec_order:02}', spec_name, 'b', 'Bảng tính' + os.path.splitext(filepath)[1]])
		else:
			os.makedirs(self.path + 'Data\\', exist_ok=True)
			try:
				_current_max_subchar = max('b', self.max_subchar[self.spec_orders[spec_name]])
			except KeyError:
				_current_max_subchar = 'b'
			_subchar = bytes([bytes(_current_max_subchar, 'utf-8')[0] + 1]).decode('utf-8')
			_new_file = 'Data\\' + '. '.join([f'{spec_order:02}', spec_name, _subchar, os.path.split(filepath)[1]])
		
		if (overwrite == False and not os.path.exists(self.path + _new_file)) \
			or (overwrite == True):
			shutil.copy2(filepath, self.path + _new_file)
			self.analyze_folder()
		
		if del_src == True:
			try:
				os.remove(filepath)
			except Exception as e:
				logger.error('Failed to remove file: ' + str(e))
		
	def add_general_doc(self, filepath, doc_type, del_src=False, overwrite=False):
		"""Add a summary file to main folder
		Rename file to defined format
		If doc_type = 'ar', the file considered as an analytical result, file named as '_<analytical_number> <product_name> Phiếu phân tích'
		If doc_type = 'cw', the file considered as a control worksheet, file named as '_<analytical_number> <product_name> Kiểm soát mẫu phân tích'
		"""
		print(filepath)
		if not os.path.isfile(filepath):
			return 1
		
		if doc_type == 'ar':
			os.makedirs(self.path, exist_ok=True)
			_new_file = '_' + self.path.split('\\')[-2] + ' ' + self.path.split('\\')[-3] + ' Phiếu phân tích' + os.path.splitext(filepath)[-1]
		elif doc_type == 'cw':	# # Subprefix = 'b'
			os.makedirs(self.path, exist_ok=True)
			_new_file = '_' + self.path.split('\\')[-2] + ' ' + self.path.split('\\')[-3] + ' Kiểm soát mẫu phân tích.' + os.path.splitext(filepath)[-1]
	
		if (overwrite == False and not os.path.exists(self.path + _new_file)) \
			or (overwrite == True):
			shutil.copy2(filepath, self.path + _new_file)
			self.analyze_folder()
	
		if del_src == True:
			if overwrite == True or not os.path.exists(self.path + _new_file):
				try:
					os.remove(filepath)
				except Exception as e:
					logger.error('Failed to remove file: ' + str(e))
