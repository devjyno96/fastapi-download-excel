import os
import requests
import json
from dotenv import load_dotenv

load_dotenv()

host = os.getenv('OPENSEARCH_HOST')
auth = (os.getenv('OPENSEARCH_USER'), os.getenv('OPENSEARCH_PASSWORD'))

index_name = 'opensearch-test'


def get_search_result(company_name, start_time, server_ip):
    query = {
        "size": 10000,
        "query": {
            "bool": {
                "must": [
                    {
                        "range": {
                            "Time": {
                                "gte": start_time
                            }
                        }
                    },
                    {
                        "match": {
                            "To_IP": server_ip
                        }
                    },
                    {
                        "match": {
                            "Company_Name": company_name
                        }
                    }
                ]
            }
        }
    }
    return requests.get(f'https://{host}/{index_name}/_search', auth=auth, json=query).json()['hits']['hits']


def get_result_only_direct_access(company_name, start_time, server_ip, not_direct_access_ip_list, size: int = 10000):
    result = []
    if len(not_direct_access_ip_list) == 0:
        query = {
            "size": size,
            "query": {
                "bool": {
                    "must": [
                        {
                            "range": {
                                "Time": {
                                    "gte": start_time
                                }
                            }
                        },
                        {
                            "match": {
                                "To_IP": server_ip
                            }
                        },
                        {
                            "match": {
                                "Company_Name": company_name
                            }
                        }
                    ]
                }
            }
        }
        search_result = requests.get(f'https://{host}/{index_name}/_search', auth=auth, json=query).json()
        result += search_result['hits']['hits']

    for not_direct_access_tool in not_direct_access_ip_list:
        query = {
            "size": size,
            "query": {
                "bool": {
                    "must": [
                        {
                            "range": {
                                "Time": {
                                    "gte": start_time
                                }
                            }
                        },
                        {
                            "match": {
                                "To_IP": server_ip
                            }
                        },
                        {
                            "match": {
                                "Company_Name": company_name
                            }
                        }
                    ],
                    "must_not": [
                        {
                            "match": {
                                "From_IP": not_direct_access_tool.access_tool_ip_address_start
                            }
                        }
                    ]
                }
            }
        }
        search_result = requests.get(f'https://{host}/{index_name}/_search', auth=auth, json=query).json()
        result += search_result['hits']['hits']

    return result
