import requests
from bs4 import BeautifulSoup
s = requests.session()


def get_csrf_tokens():
    url = 'https://dribbble.com/session/new'
    req = s.get(url).text

    soup = BeautifulSoup(req, 'lxml')

    authenticity_token = soup.find_all('input', type='hidden')[1]
    authenticity_token = authenticity_token['value']
    print(type(authenticity_token), authenticity_token)

    return authenticity_token


def login(username, password):
    url = 'https://dribbble.com/session'
    authenticity_token = get_csrf_tokens()

    data = {
        'authenticity_token': authenticity_token,
        'login': username,
        'password': password,
    }

    req = s.post(url, data=data)

    return req.text


def col_links():
    login('login', 'password')
    links_arr = []

    i = 1

    while True:

        url = 'https://dribbble.com/teams?sort=trending&page=' + str(i) + '&per_page=20'
        req = s.get(url).text
        soup = BeautifulSoup(req, 'lxml')
        links = soup.findAll('a', class_='url boosted-link')
        for link in links:
            link_full = 'https://dribbble.com' + link['href']
            links_arr.append(link_full)

        end = soup.findAll('a', class_='next_page')
        if end:
            print(i)
            i += 1
        else:
            print('End of List')
            with open('result.txt', 'w') as tmp:
                tmp.write('\n'.join(links_arr))
            print(links_arr)
            break


if __name__ == '__main__':
    # login = input('Input Your login:   \n')
    # password = input('Input Your password:   \n')
    col_links()
