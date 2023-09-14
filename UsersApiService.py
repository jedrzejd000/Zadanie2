import requests

class UsersApiService:
    def getUsers(self,resultsCount):
        url = 'https://randomuser.me/api/?results={}&inc=name, gender, dob, phone, email, location'.format(resultsCount)
        response = requests.get(url)
        if response.status_code == 200:
           data = response.json()
           return data
        else:
            print(f'Błąd: Nie udało się pobrać danych. Kod odpowiedzi: {response.status_code}')
