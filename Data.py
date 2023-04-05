class Data:
    def __init__(self):
        self._opco_code = None
        self._domain_name = None
        self._transformation = None
        self._spf_policy = None
        self._dmarc_policy = None
        self._spf_count =None

    def copy(self):
        new_data = Data()
        new_data.opco_code = self.opco_code
        new_data.domain_name = self.domain_name
        new_data.transformation = self.transformation
        new_data.spf_policy = self.spf_policy
        new_data.dmarc_policy = self.dmarc_policy
        new_data.spf_count = self.spf_count
        return new_data

    @property
    def spf_count(self):
        return self._spf_count

    @spf_count.setter
    def spf_count(self, value):
        self._spf_count = value

    @property
    def opco_code(self):
        return self._opco_code

    @opco_code.setter
    def opco_code(self, value):
        self._opco_code = value

    @property
    def domain_name(self):
        return self._domain_name

    @domain_name.setter
    def domain_name(self, value):
        self._domain_name = value

    @property
    def transformation(self):
        return self._transformation

    @transformation.setter
    def transformation(self, value):
        self._transformation = value

    @property
    def spf_policy(self):
        return self._spf_policy

    @spf_policy.setter
    def spf_policy(self, value):
        self._spf_policy = value

    @property
    def dmarc_policy(self):
        return self._dmarc_policy

    @dmarc_policy.setter
    def dmarc_policy(self, value):
        self._dmarc_policy = value