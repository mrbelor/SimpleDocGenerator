from load import load_data, transform_time
from shablon import shablon

def main():
    data = load_data("./test/data_example.xls")
    data = transform_time(data, "Время")
    
    #shablon(data, "./test/notif_shablon.docx", "Уведомления.docx")
    shablon(data, "./test/letters_shablon.docx", "Сопроводительные письма.docx")
    
if __name__ == '__main__':
    main()
