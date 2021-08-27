from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

def main(filename):
    workbook = Workbook()
    sheet = workbook.active

    # Add dados a tabela do Excel
    data_rows = [
        ["Item","Title_Book", "Kindle", "Paperback"],
        [1, "livro título 01", 9.99,  15.99],
        [2, "livro título 02", 20.99, 35.99],
        [3, "livro título 03", 14.99, 25.99],
        [4, "livro título 04", 13.99, 35.99],
        [5, "livro título 05", 11.99, 25.99],
        [6, "livro título 06", 10.99, 25.99]
    ]
    
    for row in data_rows:
        sheet.append(row)
        
    # Criando a barra de Conversa
    bar_chart = BarChart()

    data = Reference(worksheet = sheet,
                    min_row = 1,
                    max_row = 10,
                    min_col = 2,
                    max_col = 4)
    bar_chart.add_data(data, titles_from_data=True)
    sheet.add_chart(bar_chart, "E2")

    workbook.save(filename)

if __name__ == "__main__":
    main("bar_chart.xlsx")