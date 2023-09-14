from ExcelCreator import ExcelCreator
from UsersApiService import UsersApiService


def main():
   usersApiService = UsersApiService()
   resultsCount=50
   data = usersApiService.getUsers(resultsCount)
   excelCreator = ExcelCreator()
   file = excelCreator.makeFile()
   sheet = excelCreator.makeSheet(file)
   startColumn=1
   startRow=1
   rowsCount=resultsCount
   titles = ['Imię', 'Nazwisko', 'Płeć', 'Data urodzenia',
             'Numer kontaktowy', 'Adres E-mail', 'Kraj pochodzenia', 'Darmowe badania',
             'Przedział wiekowy', 'Kraj pochodzenia']

   excelCreator.makeHeadersBar(startColumn,startRow,titles,sheet)
   startRow=2
   excelCreator.loadDataToRows(startColumn, startRow,rowsCount,data,sheet)
   fileName= "nowyPlik.xlsx"
   excelCreator.saveFile(file,fileName)



if __name__ == '__main__':
    main()


