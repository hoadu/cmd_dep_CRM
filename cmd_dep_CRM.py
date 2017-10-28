import xlwings as xw
import sqlite3
import os
import datetime
import pandas as pd


def create_connection(db_file):
    try:
        conn = sqlite3.connect(db_file,
                               detect_types=sqlite3.PARSE_DECLTYPES |
                               sqlite3.PARSE_COLNAMES)
        return conn
    except Error as e:
        print(e)

    return None


def create_client(conn, client):
    sql = ''' INSERT INTO clients(date_added, OKPO, name, market, branch, manager)
              VALUES(?,?,?,?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, client)
    return cur.lastrowid


def insert_a_client():
    wb = xw.Book.caller()
    date_added = wb.sheets['database'].range("A3").value.strftime('%Y-%m-%d')
    okpo = wb.sheets['database'].range("B3").value

    client_name = wb.sheets['database'].range("C3").value
    # wb.sheets[0].range("A4").value = okpo
    # wb.sheets[0].range("B4").value = client_name
    market = wb.api.ActiveSheet.OLEObjects("ComboBox1").Object.Value
    branch = wb.api.ActiveSheet.OLEObjects("ComboBox2").Object.Value
    manager = wb.api.ActiveSheet.OLEObjects("ComboBox9").Object.Value
    database = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')

    # create a database connection
    conn = create_connection(database)

    try:
        with conn:
            # create a new client
            client = (date_added, okpo, client_name, market, branch, manager)
            client_id = create_client(conn, client)
            wb.sheets['database'].range("F22").color = (146, 208, 80)
            wb.sheets['database'].range("F22").value = \
                str(datetime.datetime.now()) + \
                ": Создан клиент " + str(client_name)

    except sqlite3.IntegrityError as e:
        wb.sheets['database'].range("F22").color = (240, 100, 77)
        wb.sheets['database'].range("F22").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


def create_service(conn, service):
    sql = ''' INSERT INTO services(date_added, client, product)
              VALUES(?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, service)
    return cur.lastrowid


def insert_a_service():
    wb = xw.Book.caller()
    date_added = wb.sheets['database'].range("A10").value.strftime('%Y-%m-%d')
    client = wb.api.ActiveSheet.OLEObjects("ComboBox3").Object.Value
    product = wb.api.ActiveSheet.OLEObjects("ComboBox4").Object.Value

    database = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')

    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
            # create a new project
            service = (date_added, client, product)
            service_id = create_service(conn, service)
            wb.sheets['database'].range("F22").color = (146, 208, 80)
            wb.sheets['database'].range("F22").value = \
                str(datetime.datetime.now()) + ": Клиенту " + str(client) \
                + ' добавлен продукт ' + str(product)
    except sqlite3.IntegrityError as e:
        wb.sheets['database'].range("F22").value = \
            str(datetime.datetime.now()) + ': ' + str(e)
        wb.sheets['database'].range("F22").color = (240, 100, 77)


def create_contact(conn, contact):
    sql = ''' INSERT INTO contacts(date_added, family, name, surname, phone, email)
              VALUES(?,?,?,?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, contact)
    return cur.lastrowid


def insert_a_contact():
    wb = xw.Book.caller()
    date_added = wb.sheets['database'].range("A17").value
    family = wb.sheets['database'].range("B17").value
    name = wb.sheets['database'].range("C17").value

    if name is None:
        wb.sheets['database'].range("F22").color = (240, 100, 77)
        wb.sheets['database'].range("F22").value = \
            "Имя контакта является обязательным для заполнения!"
        return None
    else:
        name += " "

    surname = wb.sheets['database'].range("D17").value

    if family is None:
        family = ' '
    if surname is None:
        surname = ' '

    phone = wb.sheets['database'].range("E17").value
    email = wb.sheets['database'].range("F17").value

    database = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')

    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
            # create a new project
            contact = (date_added, family, name, surname, phone, email)
            contact_id = create_contact(conn, contact)
            wb.sheets['database'].range("F22").color = (146, 208, 80)
            wb.sheets['database'].range("F22").value = \
                str(datetime.datetime.now()) + ": Создан контакт " + str(name) \
                + ' ' + str(surname) + ' ' + str(family)
    except sqlite3.IntegrityError as e:
        wb.sheets['database'].range("F22").color = (240, 100, 77)
        wb.sheets['database'].range("F22").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


def create_bounded_contact(conn, bounded_contact):
    sql = ''' INSERT INTO bounded_contacts(client, contact)
              VALUES(?,?) '''
    cur = conn.cursor()
    cur.execute(sql, bounded_contact)
    return cur.lastrowid


def insert_a_bounded_contact():
    wb = xw.Book.caller()
    client = wb.api.ActiveSheet.OLEObjects("ComboBox5").Object.Value
    contact = wb.api.ActiveSheet.OLEObjects("ComboBox6").Object.Value

    database = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')

    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
            # create a bounded_contact
            bounded_contact = (client, contact)
            bounded_contact_id = create_bounded_contact(conn, bounded_contact)
            wb.sheets['database'].range("F22").color = (146, 208, 80)
            wb.sheets['database'].range("F22").value = \
                str(datetime.datetime.now()) + ": Клиенту " + str(client) + \
                ' назначен контакт ' + str(contact)
    except sqlite3.IntegrityError as e:
        wb.sheets['database'].range("F22").color = (240, 100, 77)
        wb.sheets['database'].range("F22").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


def create_bounded_status(conn, bounded_status):
    sql = ''' INSERT INTO bounded_statuses(client, status, date_added, comment)
              VALUES(?,?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, bounded_status)
    return cur.lastrowid


def insert_a_bounded_status():
    wb = xw.Book.caller()
    date_added = \
        wb.sheets['database'].range("A27").value.strftime('%Y-%m-%d %H:%M')
    client = wb.api.ActiveSheet.OLEObjects("ComboBox7").Object.Value
    status = wb.api.ActiveSheet.OLEObjects("ComboBox8").Object.Value
    comment = wb.sheets['database'].range("D27").value

    database = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')

    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
            # create a new project
            bounded_status = (client, status, date_added, comment)
            bounded_status_id = create_bounded_status(conn, bounded_status)
            wb.sheets['database'].range("F22").color = (146, 208, 80)
            wb.sheets['database'].range("F22").value = \
                str(datetime.datetime.now()) + ": Статус клиента " \
                + str(client) + ' изменен на ' + str(status)
    except sqlite3.IntegrityError as e:
        wb.sheets['database'].range("F22").color = (240, 100, 77)
        wb.sheets['database'].range("F22").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


def create_request(conn, request):
    sql = ''' INSERT INTO requests(date_added, branch, comment)
              VALUES(?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, request)
    return cur.lastrowid


def insert_a_request():
    wb = xw.Book.caller()
    date_added = wb.sheets['database'].range("D10").value.strftime('%Y-%m-%d')
    branch = wb.api.ActiveSheet.OLEObjects("ComboBox10").Object.Value
    comment = wb.sheets['database'].range("F10").value

    database = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')

    # create a database connection
    conn = create_connection(database)
    try:
        with conn:
            # create a new request
            request = (date_added, branch, comment)
            request_id = create_request(conn, request)
            wb.sheets['database'].range("F22").color = (146, 208, 80)
            wb.sheets['database'].range("F22").value = \
                str(datetime.datetime.now()) + \
                ": Создано обращение от филиала " + str(branch)
    except sqlite3.IntegrityError as e:
        wb.sheets['database'].range("F22").color = (240, 100, 77)
        wb.sheets['database'].range("F22").value = \
            str(datetime.datetime.now()) + ': ' + str(e)


def combobox(command, combo_box_name, source_cell):
    wb = xw.Book.caller()
    source = wb.sheets['source']

    # Place the database next to the Excel file
    db_file = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')

    # Database connection and creation of cursor
    con = \
        sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES |
                        sqlite3.PARSE_COLNAMES)
    cursor = con.cursor()

    # Database Query
    cursor.execute(command)

    # Write IDs and Names to hidden sheet
    source.range(source_cell).expand().clear_contents()
    source.range(source_cell).value = cursor.fetchall()

    combo = combo_box_name
    wb.api.ActiveSheet.OLEObjects(combo).Object.ListFillRange = \
        'Source!{}'.format(str(source.range(source_cell).expand().address))
    wb.api.ActiveSheet.OLEObjects(combo).Object.BoundColumn = 1
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnCount = 2
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnWidths = 0

    # Close cursor and connection
    cursor.close()
    con.close()


def count_requests():
    wb = xw.Book.caller()
    db_file = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')
    conn = create_connection(db_file, )
    cursor = conn.cursor()
    sql = '''SELECT branches.name AS 'Филиал', COUNT(requests.id)
        AS 'Обращений за период' FROM requests
        JOIN branches on branches.id = requests.branch
        WHERE requests.date_added BETWEEN ? AND ?
        GROUP BY branches.name'''
    start_date = \
        wb.sheets['branches_report'].range("A3").value.strftime('%Y-%m-%d')
    end_date = \
        wb.sheets['branches_report'].range("C3").value.strftime('%Y-%m-%d')
    query = cursor.execute(sql, [start_date, end_date])
    cols = [column[0] for column in query.description]
    data = pd.DataFrame(query.fetchall(), columns=cols)
    wb.sheets['branches_report'].range('A8:F100').clear_contents()

    return data


def generate_branches_report():
    wb = xw.Book.caller()
    db_file = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')
    conn = create_connection(db_file, )
    cursor = conn.cursor()
    sql = '''SELECT branches.name AS 'Филиал', products.name AS 'Продукт',
        statuses.name AS 'Статус',
        COUNT(bounded_statuses.status) AS 'Количество',
        MAX(bounded_statuses.date_added) AS 'Дата'
        FROM branches
        JOIN clients on clients.branch = branches.id
        JOIN services on services.client = clients.okpo
        JOIN products on products.id = services.product
        JOIN bounded_statuses on bounded_statuses.client = clients.okpo
        JOIN statuses on statuses.id = bounded_statuses.status
            AND bounded_statuses.date_added = (
            SELECT MAX(bounded_statuses.date_added) FROM bounded_statuses
            WHERE bounded_statuses.client = clients.okpo)
        WHERE (bounded_statuses.date_added BETWEEN ? AND ?)
        GROUP BY branches.name, bounded_statuses.status'''

    start_date = \
        wb.sheets['branches_report'].range("A3").value.strftime('%Y-%m-%d')
    end_date = \
        wb.sheets['branches_report'].range("C3").value.strftime('%Y-%m-%d')
    query = cursor.execute(sql, [start_date, end_date])
    cols = [column[0] for column in query.description]
    data = pd.DataFrame(query.fetchall(), columns=cols)
    wb.sheets['branches_report'].range('F8').options(index=False).value = \
        count_requests()
    wb.sheets['branches_report'].range('A8').options(index=False).value = data


def get_all_clients():
    wb = xw.Book.caller()
    db_file = os.path.join(os.path.dirname(wb.fullname), 'cmd_dep_CRM.db')
    conn = create_connection(db_file, )
    cursor = conn.cursor()
    sql = '''SELECT DISTINCT clients.okpo AS 'ОКПО',  clients.name AS 'Компания',
        markets.name AS 'Рынок', statuses.name AS 'Текущий статус',
        bounded_statuses.date_added AS 'Статус действителен с',
        (contacts.name || contacts.family) AS 'Контактное лицо',
        contacts.phone AS 'Телефон', contacts.email AS 'Email'
        FROM clients
        JOIN markets on markets.id = clients.market
        JOIN bounded_statuses on bounded_statuses.client = clients.okpo AND
            bounded_statuses.date_added = (
            SELECT bounded_statuses.date_added FROM bounded_statuses
            WHERE bounded_statuses.client = clients.okpo
            ORDER BY bounded_statuses.date_added DESC LIMIT 1)
        JOIN statuses on statuses.id = bounded_statuses.status
        JOIN bounded_contacts on bounded_contacts.client = clients.okpo
        JOIN contacts on contacts.id = bounded_contacts.contact'''

    query = cursor.execute(sql, )
    cols = [column[0] for column in query.description]
    data = pd.DataFrame(query.fetchall(), columns=cols)
    wb.sheets['clients'].range('A8:H100').clear_contents()
    wb.sheets['clients'].range('A7').options(index=False).value = data
