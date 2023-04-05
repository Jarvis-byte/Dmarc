import json
from datetime import datetime

from PolicyChange import PolicyChange


class PolicyChangeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, PolicyChange):
            return {
                "Opco Code": obj.opco_code,
                "Domain Name": obj.domain_name,
                "Old Dmarc Policy": obj.dmarc_policy_old,
                "New Dmarc Policy": obj.dmarc_policy_new,
                "Old SPF Policy": obj.spf_policy_change_old,
                "New SPF Policy": obj.spf_policy_change_new,
                "Transformation Status -> Old": obj.transformation_old,
                "Transformation Status -> New": obj.transformation_new
            }
        else:
            return super().default(obj)
