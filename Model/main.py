from load import load_data, transform_time, transform_address
from shablon import shablon

def main():
    cfg_path = "address_module/address_config.json"
    
    data = load_data("./test/data_example.xls")
    data = transform_time(data, "Время")
    data = transform_address(data, "Адрес", config_path=cfg_path)
    
    #shablon(data, "./test/letters_shablon.docx", "Уведомления.docx")
    shablon(data, "./test/notif_shablon.docx", "Deleteme.docx", config_path=cfg_path)
    
if __name__ == '__main__':
    main()
