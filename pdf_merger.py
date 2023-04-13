import collections
import tkinter
from tkinter import ttk
from tkinter import filedialog
from pypdf import PdfWriter
import os


class PDFMerger:
    # START:

    MetadataTuple = collections.namedtuple("MetadataTuple", ("category", "stringvar"))

    def init_variables(self):
        self.file_paths = []
        self.file_base_names = []
        self.menu_buttons = []
        self.menus = []
        self.menu_path_vars = []
        self.menu_base_name_vars = []
        self.from_vars = []
        self.to_vars = []
        self.page_input_frms = []
        self.from_entries = []
        self.to_entries = []
        self.number_of_rows = 0
        self.DEFAULT_NUMBER_OF_ROWS = 5
        self.DEFAULT_INPUT = "Select a file..."
        self.indices_of_default_inputs = set(range(self.DEFAULT_NUMBER_OF_ROWS))
        self.metadata = {
            # Label: (Category name in PDF metadata, StringVar)
            "Title": self.MetadataTuple("/Title", tkinter.StringVar()),
            "Description": self.MetadataTuple("/Subject", tkinter.StringVar()),
            "Author": self.MetadataTuple("/Author", tkinter.StringVar()),
            "Organisation": self.MetadataTuple("/Creator", tkinter.StringVar()),
            "Software": self.MetadataTuple("/Producer", tkinter.StringVar()),
        }

    def create_root(self):
        self.root = tkinter.Tk()
        self.root.eval("tk::PlaceWindow . center")
        self.root.title("PDF Merger")

    def make_window_resizeable(self):
        # The following line is throwing: Wm.wm_grid() got an unexpected keyword argument 'sticky'.
        # Unsure why. If it worked, it should make the app conntents stretch when the window is resized.
        # self.root.grid(sticky=tkinter.N + tkinter.S + tkinter.E + tkinter.W)

        top = self.root.winfo_toplevel()
        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

    def create_frm(self):
        self.frm = ttk.Frame(self.root, padding=10)
        self.frm.grid()

    # Creates Tcl wrappers for self.is_valid_input() and self.reset_page_input()
    def wrap_functions(self):
        self.validate_function = self.root.register(self.is_valid_input)
        self.reset_function = self.root.register(self.reset_page_input)

    def get_file_names(self):
        paths = filedialog.askopenfilenames(
            parent=self.root, title="Select PDF files", multiple=True
        )

        # Creates a list of paths, removing any non-pdf paths
        paths = [x for x in paths if os.path.splitext(x)[1].lower() == ".pdf"]

        base_names = [os.path.basename(path) for path in paths]

        self.file_paths.extend(paths)
        self.file_base_names.extend(base_names)

    # CREATE INNER FRAMES:

    def create_file_frame(self):
        self.file_frm = ttk.Frame(self.frm, padding=10)
        self.file_frm.grid(column=0, row=0)

    def create_menu_frame(self):
        self.menu_frm = ttk.Frame(self.frm, padding=10)
        self.menu_frm.grid(column=0, row=1)

    def create_metadata_frame(self):
        self.metadata_frm = ttk.Frame(self.frm, padding=10)
        self.metadata_frm.grid(column=1, row=1, sticky=tkinter.NW)

    # FRAME CONTENT:

    def display_loaded_files(self):
        for i, file in enumerate(self.file_base_names):
            ttk.Label(self.file_frm, text=file).grid(
                column=1, row=i, padx=10, sticky=tkinter.E
            )
            ttk.Button(
                self.file_frm,
                text="X",
                command=lambda index=i: self.remove_loaded_file(index),
                width=1,
            ).grid(column=2, row=i)

    def fill_file_frame(self):
        ttk.Label(self.file_frm, text="Loaded files:").grid(column=0, row=0)
        self.display_loaded_files()
        number_of_files = len(self.file_base_names)
        ttk.Button(
            self.file_frm, text="Load more files", command=self.load_more_files
        ).grid(column=1, row=number_of_files + 1, pady=(20, 0))

    def create_header(self):
        ttk.Label(self.menu_frm, text="Position").grid(column=0, row=0, pady=10)
        ttk.Label(self.menu_frm, text="File name").grid(column=1, row=0, pady=10)
        ttk.Label(self.menu_frm, text="Pages").grid(column=2, row=0, pady=10)

    def create_number(self, i):
        ttk.Label(self.menu_frm, text=i + 1).grid(column=0, row=1 + i)

    def create_menu(self, i):
        menu_path_var = tkinter.StringVar(master=self.root, name=f"menu_path_var_{i}")
        menu_base_name_var = tkinter.StringVar(
            master=self.root, name=f"menu_base_name_var_{i}"
        )
        menu_path_var.set(self.DEFAULT_INPUT)
        menu_base_name_var.set(self.DEFAULT_INPUT)

        menu_base_name_var.trace_add("write", self.updates_after_menu_selection)

        menu_button = ttk.Menubutton(
            self.menu_frm, textvariable=menu_base_name_var, width=30
        )
        menu_button.grid(column=1, row=1 + i)
        menu = tkinter.Menu(menu_button)
        menu_button["menu"] = menu

        menu.add_radiobutton(
            label=self.DEFAULT_INPUT,
            value=self.DEFAULT_INPUT,
            variable=menu_base_name_var,
        )
        for base_name in self.file_base_names:
            menu.add_radiobutton(
                label=base_name, value=base_name, variable=menu_base_name_var
            )

        self.menu_buttons.append(menu_button)
        self.menus.append(menu)
        self.menu_path_vars.append(menu_path_var)
        self.menu_base_name_vars.append(menu_base_name_var)

    def create_page_input(self, i):
        from_var = tkinter.StringVar(master=self.root)
        to_var = tkinter.StringVar(master=self.root)
        page_input = ttk.Frame(self.menu_frm)
        page_input.grid(column=2, row=1 + i)

        from_entry = ttk.Entry(
            page_input,
            name=f"from_entry_{i}",
            textvariable=from_var,
            validatecommand=(self.validate_function, "%P", "%W", "%V"),
            validate="all",
            invalidcommand=(self.reset_function, "%W", "%V"),
            width=4,
        )
        from_entry.grid(column=0, row=0)

        ttk.Label(page_input, text="to").grid(column=1, row=0)

        to_entry = ttk.Entry(
            page_input,
            name=f"to_entry_{i}",
            textvariable=to_var,
            validatecommand=(self.validate_function, "%P", "%W", "%V"),
            validate="all",
            invalidcommand=(self.reset_function, "%W", "%V"),
            width=4,
        )
        to_entry.grid(column=2, row=0)

        self.from_vars.append(from_var)
        self.to_vars.append(to_var)
        self.page_input_frms.append(page_input)
        self.from_entries.append(from_entry)
        self.to_entries.append(to_entry)

    def create_rows(self, added_rows):
        displayed_rows = len(self.menu_buttons)

        for i in range(displayed_rows, displayed_rows + added_rows):
            self.create_number(i)
            self.create_menu(i)
            self.create_page_input(i)

    def fill_menu_frame(self):
        self.create_header()

        if self.number_of_rows < self.DEFAULT_NUMBER_OF_ROWS:
            self.create_rows(self.DEFAULT_NUMBER_OF_ROWS)
            self.number_of_rows = self.DEFAULT_NUMBER_OF_ROWS
        else:
            self.create_rows(self.number_of_rows)

        self.set_disable_states()

        ttk.Button(
            self.menu_frm, text="Add 5 rows", command=lambda: self.add_new_rows(5)
        ).grid(column=0, row=100, pady=(20, 0))
        ttk.Button(self.menu_frm, text="Save new PDF", command=self.execute).grid(
            column=1, row=100, pady=(20, 0)
        )
        ttk.Button(self.menu_frm, text="Quit", command=self.root.destroy).grid(
            column=2, row=100, pady=(20, 0)
        )

    def fill_metadata_frame(self):
        ttk.Label(self.metadata_frm, text="Metadata").grid(
            row=0, column=0, columnspan=2, pady=10
        )

        self.metadata["Software"].stringvar.set("Adobe Acrobat")

        for i, (key, value) in enumerate(self.metadata.items()):
            ttk.Label(self.metadata_frm, text=key).grid(
                row=1 + i, column=0, sticky=tkinter.W
            )
            ttk.Entry(self.metadata_frm, textvariable=value.stringvar).grid(
                row=1 + i, column=1
            )

    # CHANGES TO CONTENT:

    def reload_file_frame(self):
        self.file_frm.destroy()
        self.create_file_frame()
        self.fill_file_frame()

    def reload_menu_frame(self):
        self.menu_frm.destroy()

        self.menu_buttons = []
        self.menus = []
        self.menu_path_vars = []
        self.menu_base_name_vars = []
        self.from_vars = []
        self.to_vars = []
        self.from_entries = []
        self.to_entries = []

        self.create_menu_frame()
        self.fill_menu_frame()

    def remove_loaded_file(self, i):
        self.file_paths.pop(i)
        self.file_base_names.pop(i)

        self.reload_file_frame()
        self.reload_menu_frame()

    def load_more_files(self):
        self.get_file_names()

        self.reload_file_frame()
        self.reload_menu_frame()

    def add_new_rows(self, n):
        self.create_rows(n)
        self.number_of_rows += n
        self.set_disable_states()

    # AUTOMATIC UPDATES:

    def update_indices_of_default_inputs(self, base_name, var_index):
        if base_name == self.DEFAULT_INPUT:
            self.indices_of_default_inputs.add(var_index)
        else:
            self.indices_of_default_inputs.discard(var_index)

    def update_menu_path_var_to_match_menu_base_name_var(self, base_name, var_index):
        if base_name == self.DEFAULT_INPUT:
            path = self.DEFAULT_INPUT
        else:
            list_index = self.file_base_names.index(base_name)
            path = self.file_paths[list_index]

        self.menu_path_vars[var_index].set(path)

    def set_disable_states(self):
        zipped = zip(
            self.menu_base_name_vars,
            self.menu_buttons,
            self.from_entries,
            self.to_entries,
            strict=True,
        )

        default_input_encountered = False

        for menu_base_name_var, menu_button, from_entry, to_entry in zipped:
            if not default_input_encountered:
                menu_button.state(["!disabled"])
            else:
                menu_button.state(["disabled"])

            if menu_base_name_var.get() == self.DEFAULT_INPUT:
                default_input_encountered = True

            if not default_input_encountered:
                from_entry.state(["!disabled"])
                to_entry.state(["!disabled"])
            else:
                from_entry.state(["disabled"])
                to_entry.state(["disabled"])

    def extract_base_name_and_var_index_from_stringvar_name(self, name1):
        var_index = name1.split("_")[-1]
        var_index = int(var_index)
        base_name = self.menu_base_name_vars[var_index].get()
        return base_name, var_index

    def updates_after_menu_selection(self, name1, name2, ops):
        base_name, var_index = self.extract_base_name_and_var_index_from_stringvar_name(
            name1
        )

        self.update_indices_of_default_inputs(base_name, var_index)
        self.update_menu_path_var_to_match_menu_base_name_var(base_name, var_index)
        self.set_disable_states()

        is_default_input = var_index in self.indices_of_default_inputs
        default_input_exists_above = (
            min(self.indices_of_default_inputs, default=9999) < var_index
        )
        is_second_or_later_default_input = (
            is_default_input and default_input_exists_above
        )

        if is_default_input:
            self.from_vars[var_index].set("")
            self.to_vars[var_index].set("")

        # Avoids recursion caused by updates which happen in the following function
        if is_second_or_later_default_input:
            return

        self.reset_menus_under_first_default_input()

    def reset_menus_under_first_default_input(self):
        if not self.indices_of_default_inputs:
            return

        first_default_input_index = min(self.indices_of_default_inputs)

        # Early return if the only default input in on the last line
        if first_default_input_index == self.number_of_rows - 1:
            return

        for i in range(first_default_input_index + 1, self.number_of_rows):
            if i in self.indices_of_default_inputs:
                continue
            self.menu_base_name_vars[i].set(self.DEFAULT_INPUT)

    # CALLBACKS AND VALIDATIONS:

    def is_ascii_digit(self, new_text):
        for character in new_text:
            if ord(character) not in range(48, 58):
                return False
        return True

    def is_larger_than_zero(self, new_text):
        if int(new_text) > 0:
            return True
        else:
            return False

    def to_is_not_smaller_than_from(self, new_text, line_index):
        from_value = self.from_vars[line_index].get()
        if int(new_text) < int(from_value):
            return False
        else:
            return True

    def is_valid_input(self, new_text, widget_name, callback_reason):
        if new_text == "":
            return True
        if callback_reason == "key":
            return self.is_ascii_digit(new_text)
        if "focus" in callback_reason:
            if "from" in widget_name:
                return self.is_larger_than_zero(new_text)
            elif "to" in widget_name:
                line_index = widget_name.split("_")[2]
                line_index = int(line_index)
                return self.is_larger_than_zero and self.to_is_not_smaller_than_from(
                    new_text, line_index
                )

    def reset_page_input(self, widget_name, callback_reason):
        if "focus" not in callback_reason:
            return
        index = widget_name.split("_")[2]
        index = int(index)
        if "from" in widget_name:
            self.from_vars[index].set("")
        elif "to" in widget_name:
            self.to_vars[index].set("")

    # INPUT HANDLING:

    def extract_input(self):
        menu_path_vars_got = [x.get() for x in self.menu_path_vars]
        from_vars_got = [
            int(x.get()) if len(x.get()) > 0 else "" for x in self.from_vars
        ]
        to_vars_got = [int(x.get()) if len(x.get()) > 0 else "" for x in self.to_vars]
        try:
            zipped = zip(menu_path_vars_got, from_vars_got, to_vars_got, strict=True)
        except ValueError:
            print("[ERROR] Zipped lists aren't all the same length.")
            quit()
        self.extracted_input = list(zipped)

    def rstrip_input(self):
        first_empty_input = []
        for tup in self.extracted_input:
            if tup[0] == self.DEFAULT_INPUT:
                first_empty_input.append(tup)
                break
        if len(first_empty_input) > 0:
            first_default_input_position = self.extracted_input.index(
                first_empty_input[0]
            )
            self.clean_input = self.extracted_input[:first_default_input_position]
        else:
            self.clean_input = None  # What if None?
        del self.extracted_input

    def translate_input(self):
        self.translated_input = [
            (file_path, (from_page - 1, to_page))
            for (file_path, from_page, to_page) in self.clean_input
        ]
        del self.clean_input

    def extract_metadata_input(self):
        return {tup.category: tup.stringvar.get() for tup in self.metadata.values()}

    # CREATING THE OUTPUT FILE:

    def init_merger(self):
        self.merger = PdfWriter()

    def merge(self):
        for tup in self.translated_input:
            file_handle = open(tup[0], "rb")
            pages = tup[1]
            self.merger.append(fileobj=file_handle, pages=pages)
            file_handle.close()

    def add_metadata(self, metadata_input):
        self.merger.add_metadata(metadata_input)

    def get_output_file_name(self):
        return filedialog.asksaveasfilename(defaultextension="pdf")

    def write(self, output_file_name):
        self.output_handle = open(output_file_name, "wb")
        self.merger.write(self.output_handle)

    def close_merger(self):
        self.merger.close()
        self.output_handle.close()

    def execute(self):
        self.extract_input()
        self.rstrip_input()
        self.translate_input()
        self.init_merger()
        self.merge()
        metadata_input = self.extract_metadata_input()
        self.add_metadata(metadata_input)
        output_file_name = self.get_output_file_name()
        self.write(output_file_name)
        self.close_merger()

    # MAIN:

    def main(self):
        self.create_root()
        self.init_variables()
        self.make_window_resizeable()
        self.create_frm()
        self.wrap_functions()
        self.get_file_names()
        self.create_file_frame()
        self.fill_file_frame()
        self.create_menu_frame()
        self.fill_menu_frame()
        self.create_metadata_frame()
        self.fill_metadata_frame()

        self.root.mainloop()
