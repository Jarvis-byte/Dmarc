class PolicyChange:
    def __init__(self):
        self._opco_code = None
        self._domain_name = None
        self._dmarc_policy_old = None
        self._dmarc_policy_new = None
        self._spf_policy_change_old = None
        self._spf_policy_change_new = None
        self._transformation_old = None
        self._transformation_new = None

    # ------------------------

    @property
    def opco_code(self):
        return self._opco_code

    @opco_code.setter
    def opco_code(self, value):
        self._opco_code = value

    # ------------------------

    @property
    def domain_name(self):
        return self._domain_name

    @domain_name.setter
    def domain_name(self, value):
        self._domain_name = value

    # ------------------------
    @property
    def dmarc_policy_old(self):
        return self._dmarc_policy_old

    @dmarc_policy_old.setter
    def dmarc_policy_old(self, value):
        self._dmarc_policy_old = value

    # ------------------------

    @property
    def dmarc_policy_new(self):
        return self._dmarc_policy_new

    @dmarc_policy_new.setter
    def dmarc_policy_new(self, value):
        self._dmarc_policy_new = value

    # ------------------------

    @property
    def spf_policy_change_old(self):
        return self._spf_policy_change_old

    @spf_policy_change_old.setter
    def spf_policy_change_old(self, value):
        self._spf_policy_change_old = value

    # ------------------------
    @property
    def spf_policy_change_new(self):
        return self._spf_policy_change_new

    @spf_policy_change_new.setter
    def spf_policy_change_new(self, value):
        self._spf_policy_change_new = value

    # ------------------------

    @property
    def transformation_old(self):
        return self._transformation_old

    @transformation_old.setter
    def transformation_old(self, value):
        self._transformation_old = value

    # ------------------------

    @property
    def transformation_new(self):
        return self._transformation_new

    @transformation_new.setter
    def transformation_new(self, value):
        self._transformation_new = value
