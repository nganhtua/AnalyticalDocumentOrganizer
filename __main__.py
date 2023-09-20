import os, shutil
import pprint
import tkinter as tk
from ctypes import windll
from tkinter import ttk
import tkinterDnD  # Importing the tkinterDnD module
from winsort import winsort
from analyt_dir import AnalytDir

root_dir = "Q:\\Q10. QC-2023\\Q16. KET QUA KN TU 2023\\4. Mẫu nghiên cứu"
d = AnalytDir('')

def split_filelist(filelist):
	if filelist.count(":") == 0:
		return []
	elif filelist.count(":") == 1:
		if filelist.startswith("{"):
			return [filelist[1:-1]] + split_filelist("")
		else:
			return [filelist] + split_filelist("")
	else:
		first_colon = filelist.find(":")
		second_colon = filelist[filelist.find(":")+1:].find(":") + first_colon + 1
		sep_space = filelist[:second_colon].rfind(" ")
		first_element = filelist[:sep_space]
		remaining_elements = filelist[sep_space+1:]
		return split_filelist(first_element) + split_filelist(remaining_elements)


# You have to use the tkinterDnD.Tk object for super easy initialization,
# and to be able to use the main window as a dnd widget

# GUI LAYOUT
root = tkinterDnD.Tk()
#windll.shcore.SetProcessDpiAwareness(1)
root.title("Thêm file dữ liệu")
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)
root.minsize(800, 0)

frm_top = tk.Frame(root)
frm_top.grid(sticky="enws", row=0, column=0, padx=10, pady=10)
frm_top.grid_columnconfigure(1, weight=1)
frm_top.grid_rowconfigure(6, weight=1)

tk.Label(frm_top, text="Loại sản phẩm: ").grid(sticky="w", row=0, column=0)
cbbx_product_type = ttk.Combobox(frm_top, state="readonly", width=30, 
							  values = ["Mẫu nghiên cứu"])
cbbx_product_type.current(0)
cbbx_product_type.grid(sticky="w", row=0, column=1)

tk.Label(frm_top, text="Tên sản phẩm: ").grid(sticky="w", row=1, column=0)
cbbx_product_name = ttk.Combobox(frm_top, width=30)
cbbx_product_name.grid(sticky="w", row=1, column=1)

tk.Label(frm_top, text="Số kiểm nghiệm: ").grid(sticky="w", row=2, column=0)
cbbx_analytic_no = ttk.Combobox(frm_top, width=30, values = [""])
cbbx_analytic_no.grid(sticky="w", row=2, column=1)

tk.Label(frm_top, text="Chỉ tiêu kiểm nghiệm: ").grid(sticky="w", row=3, column=0)
cbbx_spec = ttk.Combobox(frm_top, width=30, values = [""])
cbbx_spec.grid(sticky="w", row=3, column=1)

tk.Label(frm_top, text="File hồ sơ phân tích: ").grid(sticky="w", row=4, column=0)

var_doc_path = tk.StringVar()
ent_doc_path = tk.Entry(frm_top, textvariable=var_doc_path)
ent_doc_path.grid(sticky="ew", row=4, column=1)

#btn_browse_doc = ttk.Button(frm_top, text="", width=3)
#btn_browse_doc.grid(sticky="e", row=6, column=2)

tk.Label(frm_top, text="File bảng tính: ").grid(sticky="w", row=5, column=0)

var_sprdsheet_path = tk.StringVar()
ent_sprdsheet_path = tk.Entry(frm_top, textvariable=var_sprdsheet_path)
ent_sprdsheet_path.grid(sticky="ew", row=5, column=1)

#btn_browse_sprdsheet = ttk.Button(frm_top, text="", width=3)
#btn_browse_sprdsheet.grid(sticky="e", row=7, column=2)

tk.Label(frm_top, anchor="w", justify="left", text="Dữ liệu \n(SKĐ, cân,...): ").grid(sticky="nw", row=6, column=0)

lb_extra_path = tk.Listbox(frm_top)
lb_extra_path.grid(sticky="enws", row=6, column=1)

#btn_browse_extra_data = ttk.Button(frm_top, text="", width=3)
#btn_browse_extra_data.grid(sticky="en", row=8, column=2)

tk.Label(frm_top, text="Hồ sơ chung cho 1 số kiểm nghiệm (nếu có):", font=('Segoe UI', 9, 'bold', 'underline')).grid(sticky="w", row=7, column=0, columnspan=3)

tk.Label(frm_top, text="File phiếu phân tích: ").grid(sticky="w", row=8, column=0)

var_summary_path = tk.StringVar()
ent_summary_path = tk.Entry(frm_top, textvariable=var_summary_path)
ent_summary_path.grid(sticky="ew", row=8, column=1)

#btn_browse_summary = ttk.Button(frm_top, text="", width=3)
#btn_browse_summary.grid(sticky="e", row=3, column=2)

tk.Label(frm_top, text="File phiếu kiểm soát: ").grid(sticky="w", row=9, column=0)

var_control_path = tk.StringVar()
ent_control_path = tk.Entry(frm_top, textvariable=var_control_path)
ent_control_path.grid(sticky="ew", row=9, column=1)

#btn_browse_control = ttk.Button(frm_top, text="", width=3)
#btn_browse_control.grid(sticky="e", row=4, column=2)

frm_options = tk.Frame(frm_top)
frm_options.grid(sticky="ew", row=10, column=1)
frm_options.grid_columnconfigure(1, weight=1)

var_del_src = tk.IntVar()
ckbx_del_src = ttk.Checkbutton(frm_options, text="Xóa file nguồn sau khi copy", variable = var_del_src)
ckbx_del_src.grid(sticky="w", row=0, column=0)
var_del_src.set(1)

lbl_file_prg = tk.Label(frm_options)
lbl_file_prg.grid(sticky="e", row=0, column=1)
lbl_file_prg["text"] = "0/0"

var_prg = tk.IntVar()
prgbar_copy = ttk.Progressbar(frm_options, variable = var_prg)
prgbar_copy.grid(sticky="e", row=0, column=2)
var_prg.set(0)

frm_bottom = tk.Frame(root)
frm_bottom.grid(row=1, column=0, padx=10, pady=5)
frm_bottom.grid_columnconfigure(1, weight=1)

btn_save = ttk.Button(frm_bottom, text="Lưu file")
btn_save.grid(row=0, column=0, padx=10)

btn_clear = ttk.Button(frm_bottom, text="Nhập lại")
btn_clear.grid(row=0, column=1, padx=10)

btn_help = ttk.Button(frm_bottom, text="Hướng dẫn")
btn_help.grid(row=0, column=2, padx=10)

btn_exit = ttk.Button(frm_bottom, text="Thoát")
btn_exit.grid(row=0, column=3, padx=10)

# UI FUNCTIONS
def list_subdir(path):
	subfolders_w_path = [f.path for f in os.scandir(path) if f.is_dir()]
	return winsort([f.split("\\")[-1] for f in subfolders_w_path])

def on_select_product_name(event):
	w = event.widget
	value = w.get().strip()
	cbbx_product_name.set(value)
	if value != '':
		cbbx_analytic_no["values"] = list_subdir(root_dir + "\\" + value)
	cbbx_analytic_no.set("")
	cbbx_spec.set("")
	
def on_select_analytic_no(event):
	global d
	cbbx_product_name.set(cbbx_product_name.get().strip())
	product_name = cbbx_product_name.get()
	analytic_no = event.widget.get().strip()
	analytic_no_dir = root_dir + "\\" + product_name + "\\" + analytic_no
	cbbx_spec["values"] = []
	d = AnalytDir(analytic_no_dir)
	print('d.path: ' + d.path)
	if product_name != '':
		cbbx_spec["values"] = list(d.spec_orders.keys())
	cbbx_spec.set("")

def drop_files(event):
	# This function is called, when stuff is dropped into a widget
	filenames = split_filelist(event.data)
	if isinstance(event.widget, tk.Entry):
		varname = event.widget.cget("textvariable")
		event.widget.setvar(varname, filenames[0])
	else:
		for filename in filenames:
			event.widget.insert('end', filename)

def update_progress(done, total):
	if total == 0:
		lbl_file_prg["text"] = "0/0"
		var_prg.set(0)
	else:
		lbl_file_prg["text"] = "/".join([str(done), str(total)])
		var_prg.set(int(done/total*100))

def save_files(event):
	# This function called when the users click Lưu file button
	global cbbx_product_name, cbbx_analytic_no, cbbx_spec, \
		var_doc_path, var_sprdsheet_path, frm_extra_data, d
	summary_fpath = var_summary_path.get()
	control_fpath = var_control_path.get()
	doc_fpath = var_doc_path.get()
	sprdsheet_fpath = var_sprdsheet_path.get()
	product_name = cbbx_product_name.get()
	analytic_no = cbbx_analytic_no.get()
	spec_name = cbbx_spec.get()
	analytic_no_dir = root_dir + "\\" + product_name + "\\" + analytic_no
	d = AnalytDir(analytic_no_dir)
	files_done = 0
	files_count = (summary_fpath != "") + (control_fpath != "") + (doc_fpath != "") + \
					(sprdsheet_fpath != "") + lb_extra_path.index(tk.END)
	update_progress(files_done, files_count)
	if summary_fpath != "":
		d.add_general_doc(summary_fpath, "ar", del_src=var_del_src.get())
		files_done += 1
		var_summary_path.set("")
		update_progress(files_done, files_count)
	if control_fpath != "":
		d.add_general_doc(control_fpath, "cw", del_src=var_del_src.get())
		files_done += 1
		var_control_path.set("")
		update_progress(files_done, files_count)
	if product_name == '':
		print('Error: Blank product name.')
		return None
	if analytic_no == '':
		print('Error: Blank analytical number.')
		return None
	if spec_name == '':
		print('Error: Blank spec name.')
		return None
	
	if doc_fpath != "":
		d.add_file(doc_fpath, cbbx_spec.get().strip(), 'doc', del_src=var_del_src.get())
		files_done += 1
		var_doc_path.set("")
		update_progress(files_done, files_count)
	if sprdsheet_fpath != "":
		d.add_file(sprdsheet_fpath, cbbx_spec.get().strip(), 'ss', del_src=var_del_src.get())
		files_done += 1
		var_sprdsheet_path.set("")
		update_progress(files_done, files_count)
	for i, listbox_entry in enumerate(lb_extra_path.get(0, tk.END)):
		d.add_file(listbox_entry, cbbx_spec.get().strip(), 'ex', del_src=var_del_src.get())
		files_done += 1
		lb_extra_path.delete(0)
		update_progress(files_done, files_count)
	cbbx_spec.set("")

def clear_form(event):
	# This function called when the users click Nhập lại button
	cbbx_product_name.set("")
	cbbx_analytic_no.set("")
	cbbx_spec.set("")
	var_doc_path.set("")
	var_sprdsheet_path.set("")
	lb_extra_path.delete(0, 'end')
	var_del_src.set(1)

def close_windows(event):
	root.destroy()

# BINDING UI COMPONENTS WITH FUNCTIONS
cbbx_product_name.bind("<<ComboboxSelected>>", on_select_product_name)
cbbx_product_name.bind("<FocusOut>", on_select_product_name)
cbbx_analytic_no.bind("<<ComboboxSelected>>", on_select_analytic_no)
cbbx_analytic_no.bind("<FocusOut>", on_select_analytic_no)

ent_summary_path.register_drop_target("*")
ent_summary_path.bind("<<Drop>>", drop_files)

ent_control_path.register_drop_target("*")
ent_control_path.bind("<<Drop>>", drop_files)

ent_doc_path.register_drop_target("*")
ent_doc_path.bind("<<Drop>>", drop_files)

ent_sprdsheet_path.register_drop_target("*")
ent_sprdsheet_path.bind("<<Drop>>", drop_files)

lb_extra_path.register_drop_target("*")
lb_extra_path.bind("<<Drop>>", drop_files)

btn_save.bind("<Button-1>", save_files)
btn_clear.bind("<Button-1>", clear_form)
btn_exit.bind("<Button-1>", close_windows)

# STARTUP COMMANDS
cbbx_product_name["values"] = list_subdir(root_dir)

root.mainloop()
