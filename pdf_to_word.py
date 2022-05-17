from pdf2docx import Converter
import PySimpleGUI as sg

def pdf2word(file_path):
    file_name = file_path.split('.')[0]
    docx_file = f'{file_name}.docx'

    p2w = Converter(file_path)
    p2w.convert(docx_file, start=0, end=None)
    p2w.close()
    return docx_file
def main():
    sg.theme('BlueMono')

    layout = [
        [sg.Text('PDF轉Word',font=("蘋方-繁", 12)),
        sg.Text('',key='filename',size=(50,1),font=("蘋方-繁", 10),text_color='blue')],
        [sg.Output(size=(80, 10),font=("蘋方-繁", 10))],
        [sg.FileBrowse('選擇文件',key='file',target='filename'),sg.Button('開始轉換'),sg.Button('退出')]
    ]

    window = sg.Window('PDF to Word by ParrotXray Edit', layout,font=("蘋方-繁", 15),default_element_size=(50,1))

    while True:
        event, values = window.read()
        if event in (None, '退出'):
            break
        if event == '開始轉換':
            if values['file'] and values['file'].split('.')[1]=='pdf':
                file_path = pdf2word(values['file'])
                print('\n'+'轉換成功'+'....'+'\n')
                print('word文件位置:', file_path)
            else:
                print('未選取文件或非pdf格式\n請選擇文件')
    window.close()


main()

