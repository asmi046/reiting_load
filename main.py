from excel_loader import get_src_data, get_src_data_index
import tabulate



def main():
    file = "data/МО.xlsx"
    # to_base_data = get_src_data(file)
    to_base_data = get_src_data_index(file, index=2)
    raw = tabulate.tabulate(to_base_data["raw_data"], headers=['Наименование МО', 'Направление', 'Показатель', 'Значение', 'Балл', 'Лидер', 'Значение лидера', 'Место', 'Период'])

    print(raw)
    print(to_base_data["rank_values"])
    print(to_base_data["rank_index"])


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

