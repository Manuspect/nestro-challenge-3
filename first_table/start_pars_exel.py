import create_data_csv
import customs_duties_parser
import parcing_first_table
import parsing_brent_cost
import urals_parser

def main():
    # Вызываем функции из каждого модуля
    create_data_csv.get_datas()
    create_data_csv.get_kurses()
    create_data_csv.get_AHAJ()
    create_data_csv.get_ARAS()

    customs_duties_parser.main()

    parcing_first_table.main()

    parsing_brent_cost.main()

    urals_parser.main()

if __name__ == "__main__":
    main()
