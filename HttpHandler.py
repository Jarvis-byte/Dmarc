import time

import checkdmarc
import requests
import json

from ExportToExcel import ExportToExcel


class HttpHandler:
    def get_request(self, data_list, start, end):

        for a in range(start, end):
            # print(f'{a} -> {data_list[a].domain_name}')
            self.get_spf_policy(data_list, a)
            self.get_dmarc_policy(data_list, a)
            self.get_spf_count(data_list, a)
            # self.add(data_list, a)
            # print(f'Data list at {a} -> {data_list[a].dmarc_policy}')
        # print(data_list[43].dmarc_policy)

        return data_list[start:end]

        # result_queue.put(data_list[start:end])
        # print(f'Data list at {1} -> {data_list[1].dmarc_policy}')

    def get_dmarc_policy(self, list, a):
        domain = list[a].domain_name
        url = f"https://dns.google/resolve?name=_dmarc.{domain}&type=TXT"
        response = requests.get(url)

        if response.status_code != 200:
            list[a].dmarc_policy = "ERROR"
            return
        else:
            inline = response.text

        data_obj = json.loads(inline)
        # print(data_obj)
        if "Answer" in data_obj:

            answer = data_obj["Answer"]

            for i in range(len(answer)):
                ans_obj = answer[i]
                str1 = str(ans_obj)
                # print(f'Well {i} {str1}')
                if "v=DMARC1" in str1:
                    # print(f'{ans_obj["data"]} in {i}')
                    list[a].dmarc_policy = ans_obj["data"]

                    break
                else:
                    list[a].dmarc_policy = "No DMARC Record"
        else:
            list[a].dmarc_policy = "No DMARC Record"
        print(f'Domain is {domain} -> {list[a].dmarc_policy}')

    def get_spf_policy(self, list, a):
        domain = list[a].domain_name
        # print(domain)
        url = f"https://dns.google/resolve?name={domain}&type=TXT"
        response = requests.get(url)

        if response.status_code != 200:
            list[a].domain_name = "ERROR"
            return
        else:
            inline = response.text

        data_obj = json.loads(inline)
        if "Answer" in data_obj:
            answer = data_obj["Answer"]
            for i in range(len(answer)):
                ans_obj = answer[i]
                str1 = str(ans_obj)
                if "v=spf1" in str1:
                    list[a].spf_policy = ans_obj["data"]
                    break
                else:
                    list[a].spf_policy = "No SPF Record"
        else:
            list[a].spf_policy = "No SPF Record"

    def get_spf_count(self, list_data, a):
        values = [""]
        domain = list_data[a].domain_name

        try:
            dns_loockup = checkdmarc.get_spf_record(domain, nameservers=None, resolver=None, timeout=2.0)
            values = list(dns_loockup.values())
        except checkdmarc.SPFRecordNotFound as e:
            values[0] = "SPF record Not Found"

        except checkdmarc.SPFTooManyDNSLookups as e:
            values[0] = "SPF record has more then 10 spf record"

        except:
            values[len(values) - 1] = "Error in Fetching SPF"

        list_data[a].spf_count = str(values[0])

