class People(object):

    def __init__(self, combi_name=None, status=None, sso=None, First_Name=None, Last_Name=None, bu=None, company=None):
        self.combi_name = combi_name
        self.sso = sso
        self.First_Name = First_Name
        self.Last_Name = Last_Name
        self.bu = bu
        self.status = status
        self.company = company

    def print_people(self):
        print self.combi_name, "\t", self.status, "\t", self.company
