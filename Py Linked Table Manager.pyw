# Python 2.7
try:
    import Tkinter as tk
    import tkMessageBox as messagebox
    import tkFileDialog as filedialog
    import tkFont as font
    import ttk
except ImportError:
    # Python 3
    import tkinter as tk
    from tkinter import messagebox, filedialog, font, ttk

from PIL import ImageTk, Image, ImageOps
import os, win32com.client
from TreeDataView import TreeDataView

class Main:
    def __init__(self, parent):

        def SourceDatabase():
            self.filename = filedialog.askopenfilename(initialdir='C:\\', filetypes = (('Access Database','*.mdb;*.mda;*.mde;*.mdw'),))

            if self.filename != '' and self.filename != None:
                self.entry1.delete(0, tk.END)
                self.entry1.insert(tk.END, self.filename)
                try:
                    accessEngine = win32com.client.gencache.EnsureDispatch ('DAO.DBEngine.36')
                    db = accessEngine.OpenDatabase(str(self.filename))

                    tdv1.clear()
                    for tabledef in db.TableDefs:
                        if not str(tabledef.Name)[:4] == 'MSys':
                            tdf = db.TableDefs(tabledef.Name)
                            dbString = str(str(tdf.Connect).split('\\')[-1]).strip()
                            if dbString != '' and dbString != None:
                                tdv1.insert_data((str(tabledef.Name), str(tdf.Connect)))
                except Exception as e:
                    messagebox.showerror('Error',e)
                    try:
                        db.Close
                        db = None
                    except:
                        pass
                else:
                    db.Close
                    db = None
            else:
                pass

        def DestinationDatabase():
            self.openDir = filedialog.askdirectory(initialdir='')
            if self.openDir != '' and self.openDir != None:
                self.entry2.delete(0, tk.END)
                self.entry2.insert(tk.END, str(self.openDir))

        def OK():
            SourcePath = str(self.entry1.get().strip())
            DestinationPath = str(self.entry2.get().strip())
            if SourcePath == None or SourcePath == '' or DestinationPath == None or DestinationPath == '':
                pass
            else:
                try:
                    accessEngine = win32com.client.gencache.EnsureDispatch ('DAO.DBEngine.36')
                    db = accessEngine.OpenDatabase(SourcePath)
                except Exception as e:
                    messagebox.showerror('Error',e)
                else:
                    tbConString = []
                    selitems = tdv1.selection()
                    if selitems:
                        for i in selitems:
                            text = tdv1.item(i, 'values')
                            tbConString.append(text[0])
                    tdv1.clear()
                    try:
                        for Table in tbConString:
                            tdf = db.TableDefs(Table)
                            DBFile = str(tdf.Connect).split('\\')[-1].strip()
                            if Table != '' and Table != None:
                                if os.path.isfile(str(DestinationPath)+'/'+DBFile):
                                    tdf.Connect = ';DATABASE='+str(DestinationPath)+'/'+DBFile
                                    tdf.RefreshLink()
                                    if Table != '' and Table != None:
                                        tdv1.insert_data((Table, str(tdf.Connect)))
                                    else:
                                        'maybe here?'
                                else:
                                    try:
                                        mb = MessageBox(None,'File '+DBFile+\
                                                        ' was not found! Press OK to skip this file.',\
                                                        'Warning',flags.MB_OK | flags.MB_ICONWARNING)
                                    except Exception as e:
                                        break
                                        messagebox.showerror('Error',e)
                                        try:
                                            db.Close
                                            db = None
                                        except:
                                            pass
                            else:
                                pass
                    except Exception as e:
                        messagebox.showerror('Error',e)
                        try:
                            db.Close
                            db = None
                        except:
                            pass
                    else:
                        db.Close
                        db = None

        def SelectAll():
            for item in tdv1.get_children():
                tdv1.selection_add(item)

        def DeselectAll():
            selitems = tdv1.selection()
            tdv1.selection_remove(selitems)
        
        label1 = tk.Label(parent, text='Select the Database for the linked tables to be updated:')
        label1.place(x=12, y=12)
        label2 = tk.Label(parent, text='Select the Database that you wish to refresh the links to:')
        label2.place(x=12, rely=1.0, y=-33)

        f = tk.Frame()
        f.pack(side=tk.RIGHT, fill=tk.Y, padx=(0,12), pady=(45,45))

        button1 = ttk.Button(f, text='OK', command=OK)
        button1.pack(side=tk.TOP)

        button2 = ttk.Button(f, text='Close', command=parent.destroy)
        button2.pack(side=tk.TOP, pady=(12,0))

        button3 = ttk.Button(f, text='Select All', command=SelectAll)
        button3.pack(side=tk.TOP, pady=(12,0))

        button4 = ttk.Button(f, text='Deselect All', command=DeselectAll)
        button4.pack(side=tk.TOP, pady=(12,0))

        button5 = ttk.Button(parent, text='Browse', command=SourceDatabase)
        button5.place(x=319, y=12)

        button6 = ttk.Button(parent, text='Browse', command=DestinationDatabase)
        button6.place(x=319, rely=1.0, y=-33)

        self.entry1 = ttk.Entry(parent)
        self.entry1.place(w=273, x=407, y=12)

        self.entry2 = ttk.Entry(parent)
        self.entry2.place(w=273, x=407, rely=1.0, y=-33)

        tree_columns = ['Table','Connection String']
        tdv1 = TreeDataView(parent, tree_columns, scrollbar_x=False, scrollbar_y=True)
        tdv1.pack(fill='both', expand=1, padx=(12,12), pady=(45,45))
        tdv1.column(1, width=350)
     
def main():
    root = tk.Tk()
    root.title('Py Linked Table Manager (Access Versions 97-03)')
    w=768; h=450
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w) / 2
    y = (sh - h) / 2
    root.geometry('{0}x{1}+{2}+{3}'.format(w, h, int(x), int(y)))
    root.resizable(False, False)
    try:
        root.wm_iconbitmap('images/Creative-Freedom-Free-Vibrant-Table.ico')
    except:
        pass
    
    win = Main(root)

    root.mainloop()

if __name__ == '__main__':
    main()
