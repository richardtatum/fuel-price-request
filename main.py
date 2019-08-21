from models import EmailHandler, AirportSearch
from dotenv import load_dotenv
import os

load_dotenv()
file = os.getenv('file')
ops = os.getenv('ops')
bcc = os.getenv('bcc')

airport = AirportSearch(file)
email = EmailHandler()


def main():
    print(
    """
    Fuel Price Requester v2.0
    -------------------------

    Add airports to the list. Once ready, type \'finish\'
    to create the email.

    """
        )

    cat = input('Is this a Private or Commercial flight?: ')
    print(f'>{cat.title()} flight\n')

    while True:
        search = input('Please enter the Airport ICAO/IATA code: ')
        if search.lower() == 'finish':
            data = airport.format_list()
            email.create(ops, cat, bcc, data)
            email.show()
            exit()
        else:
            code = airport.search(search)
            if code is not None:
                agent = input('Please enter the handling agent: ')
                airport.add_handling(code, agent)
                airport.print_list()


if __name__ == '__main__':
    main()
