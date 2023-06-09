import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
import tkinter.scrolledtext as scrolledtext
from tkinter import messagebox
import xlwings as xw
import keyword


data = []

def apply_syntax_highlighting(event):
    # キーワードの色
    keyword_color = 'blue'
    # コメントの色
    comment_color = 'green'
    # ユーザー定義の色
    user_defined_color = 'red'
    
    text.tag_remove('keyword', '1.0', 'end')
    text.tag_remove('comment', '1.0', 'end')
    text.tag_remove('user_defined', '1.0', 'end')

    code = text.get('1.0', 'end-1c')
    lines = code.split('\n')
    user_defined_words = ['=', '(', ')','!','+','-','*','/','else:','break','continue']
    
    for index, line in enumerate(lines):
        words = line.split(' ')
        for word in words:
            if word in keyword.kwlist:
                start = f'{index+1}.{line.index(word)}'
                end = f'{index+1}.{line.index(word) + len(word)}'
                text.tag_add('keyword', start, end)
            elif word.startswith('#'):
                start = f'{index+1}.{line.index(word)}'
                end = f'{index+1}.{line.index(word) + len(word)}'
                text.tag_add('comment', start, end)
        
        for user_defined_word in user_defined_words:
            start_index = line.find(user_defined_word)
            while start_index >= 0:
                end_index = start_index + len(user_defined_word)
                start = f'{index+1}.{start_index}'
                end = f'{index+1}.{end_index}'
                text.tag_add('user_defined', start, end)
                start_index = line.find(user_defined_word, end_index)

    text.tag_config('keyword', foreground=keyword_color)
    text.tag_config('comment', foreground=comment_color)
    text.tag_config('user_defined', foreground=user_defined_color)


def open_file():
    filename = filedialog.askopenfilename(title='Open a File', filetype=(('Excel files', '*.xlsx'), ('All Files', '*.*')))
    if filename:
        try:
            # Excelアプリケーションを起動し、ブックを開く
            app = xw.App(visible=False)  # Excelアプリケーションを非表示にする場合はvisible=Falseに設定
            workbook = app.books.open(filename)
            fileLabel.config(text=filename+'が読み込まれました')
            get_excel(app, workbook)
        except ValueError:
            messagebox.showerror('ファイル破損エラー', 'File could not be opened')
        except FileNotFoundError:
            messagebox.showerror('Error', 'File Not Found')

def get_excel(app, workbook):
    global data
    # アクティブなシートを取得
    sheet = workbook.sheets.active

    # 使用されているセル範囲を取得
    used_range = sheet.used_range

    # セル範囲を2次元配列として取得
    data = used_range.options(ndim=2, empty='').value

    #print(data)

    # Excelアプリケーションを終了
    app.quit()

    # データの表示
    #print(data)

def execute_code():
    global data
    code = text.get('1.0', 'end-1c')
    exeLabel.config(text='実行中')
    print('プログラムが実行されました')
    try:
        local_vars = {}
        local_vars['data'] = data
        exec(code, globals(), local_vars)
        data = local_vars['data']
        exeLabel.config(text='実行完了')
    except Exception as e:
        messagebox.showerror('Pythonエラー', str(e))

def save_to_excel():
    global data
    filename = filedialog.asksaveasfilename(title='Save as Excel', defaultextension='.xlsx', filetypes=(('Excel files', '*.xlsx'), ('All Files', '*.*')))
    if filename:
        try:
            # Excelアプリケーションを起動し、ブックを新規作成
            app = xw.App(visible=False)  # Excelアプリケーションを非表示にする場合はvisible=Falseに設定
            workbook = app.books.add()
            sheet = workbook.sheets.active
            #print(data)
            xw.Range("A1").options(ndim=2, empty='').value = data
            # ブックを保存
            workbook.save(filename)

            # Excelアプリケーションを終了
            app.quit()

            messagebox.showinfo('Success', 'データが正しく格納されました')
        except Exception as e:
            messagebox.showerror('エクセル出力時エラー', str(e))


root = tk.Tk()
root.title('excelPlus')
root.state('zoomed')

m = tk.Menu(root)
root.config(menu=m)
file_menu = tk.Menu(m, tearoff=False)
m.add_cascade(label='メニュー', menu=file_menu)
file_menu.add_command(label='エクセルを読み込む', command=open_file)
file_menu.add_command(label='実行結果を保存', command=save_to_excel)

fileLabel = tk.Label(text='ファイルが読み込まれていません')
fileLabel.pack()

# コード入力欄
text = scrolledtext.ScrolledText(root, height=10)
text.insert('1.0','#ここではpythonが記述でき、読み込んだエクセルは2次元配列dataとして格納されています。メニューからデータを選択して、print(data)してみましょう')
apply_syntax_highlighting(0)
text.bind('<KeyRelease>', apply_syntax_highlighting)
text.pack(fill=BOTH, expand=True)

# Create a button to execute the code
exeLabel = tk.Label(text='')
exeLabel.pack()
execute_button = tk.Button(root, text='▶実行', command=execute_code)
execute_button.pack()

root.mainloop()
