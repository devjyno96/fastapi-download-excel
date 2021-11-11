from fastapi.responses import FileResponse
from FastapiDownloadExcel.opensearch_util import get_search_result

from openpyxl import Workbook
from openpyxl.styles import Font


def download_excel():
    wb = Workbook()

    ws = wb.active
    column_list = ['request_time',
                   'server_name',
                   'server_ip',
                   'table_name',
                   'table_group_name_list',
                   'query',
                   'query_type',
                   'agent_tool_name',
                   'start_datetime',
                   'query_executor_ip',
                   'query_executor',
                   ]

    data = get_search_result_repo()
    ws.append(['  '] + column_list)
    ws.column_dimensions['A'].width = 6  # INDEX
    ws.column_dimensions['B'].width = 25  # request_time
    ws.column_dimensions['C'].width = 16  # server_name
    ws.column_dimensions['D'].width = 16  # server_ip
    ws.column_dimensions['E'].width = 20  # table_name
    ws.column_dimensions['F'].width = 30  # table_group_name_list
    ws.column_dimensions['G'].width = 100  # query
    ws.column_dimensions['H'].width = 16  # query_type
    ws.column_dimensions['I'].width = 30  # agent_tool_name
    ws.column_dimensions['J'].width = 30  # start_datetime
    ws.column_dimensions['K'].width = 30  # query_executor_ip
    ws.column_dimensions['L'].width = 30  # query_executor

    ws.auto_filter.ref = 'B1:L1'

    for cell in ws["1:1"]:
        cell.font = Font(color="00000000", bold=True, size=15)

    ws.freeze_panes = 'B2'  # index, title 고정

    for index, item in enumerate(data):
        ws.append([index + 1] + list(item.values()))
    dest_filename = '/Users/logstack/Desktop/Project/fastapi-download-excel/empty_book.xlsx'
    wb.save(dest_filename)
    return FileResponse(path='/Users/logstack/Desktop/Project/fastapi-download-excel/empty_book.xlsx',
                        filename='empty_book.xlsx')


def get_search_result_repo():
    company_server_list = [
        {
            "agent_server_name": "server1",
            "agent_server_ip_address": "172.31.39.118",
            "agent_server_connection_status": True,
            "agent_server_health_status": True,
            "os_version": "Linux",
            "agent_server_id": 5
        },
        {
            "agent_server_name": "server2",
            "agent_server_ip_address": "172.31.42.21",
            "agent_server_connection_status": True,
            "agent_server_health_status": True,
            "os_version": "Linux",
            "agent_server_id": 6
        }
    ]
    company_table_list = [
        {
            "agent_table_name": "auth_group",
            "db_platform_name": "MySQL",
            "agent_table_id": 17,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_group",
            "db_platform_name": "MySQL",
            "agent_table_id": 18,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_group_permissions",
            "db_platform_name": "MySQL",
            "agent_table_id": 19,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_group_permissions",
            "db_platform_name": "MySQL",
            "agent_table_id": 20,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_permission",
            "db_platform_name": "MySQL",
            "agent_table_id": 21,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_permission",
            "db_platform_name": "MySQL",
            "agent_table_id": 22,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_user",
            "db_platform_name": "MySQL",
            "agent_table_id": 23,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_user",
            "db_platform_name": "MySQL",
            "agent_table_id": 24,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_user_groups",
            "db_platform_name": "MySQL",
            "agent_table_id": 25,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_user_groups",
            "db_platform_name": "MySQL",
            "agent_table_id": 26,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_user_user_permissions",
            "db_platform_name": "MySQL",
            "agent_table_id": 27,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "auth_user_user_permissions",
            "db_platform_name": "MySQL",
            "agent_table_id": 28,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "django_admin_log",
            "db_platform_name": "MySQL",
            "agent_table_id": 29,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "django_admin_log",
            "db_platform_name": "MySQL",
            "agent_table_id": 30,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "django_content_type",
            "db_platform_name": "MySQL",
            "agent_table_id": 31,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "django_content_type",
            "db_platform_name": "MySQL",
            "agent_table_id": 32,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "django_migrations",
            "db_platform_name": "MySQL",
            "agent_table_id": 33,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "django_migrations",
            "db_platform_name": "MySQL",
            "agent_table_id": 34,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "django_session",
            "db_platform_name": "MySQL",
            "agent_table_id": 35,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "django_session",
            "db_platform_name": "MySQL",
            "agent_table_id": 36,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "mysql_test_table1",
            "db_platform_name": "MySQL",
            "agent_table_id": 37,
            "agent_server": {
                "agent_server_name": "server1",
                "agent_server_ip_address": "172.31.39.118",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 5
            },
            "agent_table_groups": []
        },
        {
            "agent_table_name": "mysql_test_table1",
            "db_platform_name": "MySQL",
            "agent_table_id": 38,
            "agent_server": {
                "agent_server_name": "server2",
                "agent_server_ip_address": "172.31.42.21",
                "agent_server_connection_status": True,
                "agent_server_health_status": True,
                "os_version": "Linux",
                "agent_server_id": 6
            },
            "agent_table_groups": []
        }
    ]
    company_access_tool_list = [{
        "access_tool_ip_address_start": "13.125.135.221",
        "access_tool_ip_address_end": "13.125.135.221",
        "query_executor": "query_executor_test_1",
        "access_tool_id": 1,
        "agent_server": {
            "agent_server_name": "server1",
            "agent_server_ip_address": "172.31.39.118",
            "agent_server_connection_status": True,
            "agent_server_health_status": True,
            "os_version": "Linux",
            "agent_server_id": 5
        },
        "agent_tool": {
            "agent_tool_name": "Django",
            "is_direct_access": False,
            "agent_tool_id": 4
        },
        "modified_datetime": None
    }]

    company_name = 'LogStack'
    start_time = '2021-11-09T08:00:00.000000'
    end_time = '2021-11-09T09:00:00.000000'
    server_ip = '172.31.39.118'
    search_result = get_search_result(company_name=company_name,
                                      start_time=start_time,
                                      end_time=end_time,
                                      server_ip=server_ip)

    def get_server_by_ip(search_ip):
        for server in company_server_list:
            if search_ip == server['agent_server_ip_address']:
                return server

    def get_table_in_query(query_in_data):
        for table in company_table_list:
            if table['agent_table_name'] in query_in_data:
                return table

    def get_access_tool_by_ip(search_ip):
        for access_tool in company_access_tool_list:
            if access_tool['access_tool_ip_address_start'] <= search_ip <= access_tool['access_tool_ip_address_end']:
                return access_tool

    data_to_excel = []
    for row in search_result:
        row_data = row['_source']
        search_server = get_server_by_ip(row_data['To_IP'])
        search_table = get_table_in_query(row_data['MYSQL_QUERY'])
        search_access_tool = get_access_tool_by_ip(row_data['From_IP'])

        search_table_group_name = ''
        if search_table is None or search_table['agent_table_groups'] is None:
            search_table_name = ''
            search_table_group_name = ''
        else:
            for item in search_table['agent_table_groups']:
                search_access_tool += item['agent_table_group_name'] + ', '
            search_table_group_name = search_table_group_name[:-2]
            search_table_name = search_table['agent_table_name']

        request_time = row_data['Time']
        server_name = search_server["agent_server_name"]
        table_group_name_list = search_table_group_name
        query = row_data['MYSQL_QUERY']
        query_type = row_data['QUERY_TYPE']
        agent_tool_name = search_access_tool['agent_tool']['agent_tool_name']
        start_datetime = 'Not implemented'
        query_executor = search_access_tool['query_executor']

        data_to_excel.append({
            'request_time': request_time,
            'server_name': server_name,
            'server_ip': row_data['To_IP'],
            'table_name': search_table_name,
            'table_group_name_list': table_group_name_list,
            'query': query,
            'query_type': query_type,
            'agent_tool_name': agent_tool_name,
            'start_datetime': start_datetime,
            'query_executor_ip': row_data['From_IP'],
            'query_executor': query_executor,
        })
    return data_to_excel
