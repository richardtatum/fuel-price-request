import win32com.client
import csv


class EmailHandler:
    def __init__(self):
        self.olMailItem = 0x0
        self.obj = win32com.client.Dispatch("Outlook.Application")

    def create(self, to, category, bcc, format_list):
        self.new_mail = self.obj.CreateItem(self.olMailItem)
        self.new_mail.To = to
        self.new_mail.Subject = f'Fuel Price Request | {category.title()} flight'
        self.new_mail.BCC = bcc
        self.new_mail.BodyFormat = 2
        self.new_mail.HTMLBody = f"""\
                <HTML>
                    <body>
                        <span style="color:black;font-size:11pt;font-family:calibri, \
                            sans-serif">
                            <p> Dear Ops,<br>
                                <br>
                                Kindly request fuel prices for the following \
                                airports and handling agents:<br>
                                <br>
                                {format_list}<br>
                                <br>
                                I look forward to your reply.<br>
                                <br>
                                Kind regards,<br>
                                EJC Ops<br>
                            </span>
                        </p>
                    </body>
                </html>
                """

    def recipients(self, email):
        pass

    def show(self):
        self.new_mail.display()

    def send(self):
        self.new_mail.Send()


class AirportSearch:
    def __init__(self, file):
        self.file = file
        self.airport_list = []

    def search(self, code):
        with open(self.file, encoding='Latin-1') as f:
            reader = csv.reader(f, delimiter=',')
            for row in reader:
                if code.lower() in [row[10].lower(), row[0].lower()]:
                    airport_code = f'{row[0]}/{row[10]}'
                    print(f'>{airport_code}\n')
                    break
        try:
            return airport_code
        except UnboundLocalError:
            print('Airport not found. Please enter a correct IATA/ICAO code.')
            return None

    def add_handling(self, code, agent):
        completed = f'{code} - {agent}'
        print(completed)
        print()
        self.airport_list.append(completed)

    def print_list(self):
        print('Current Airport List')
        print('-------------------------')
        for airport in self.airport_list:
            print(airport)
        print()

    def format_list(self):
        self.formatted = '<br>'.join(self.airport_list)
        return self.formatted
