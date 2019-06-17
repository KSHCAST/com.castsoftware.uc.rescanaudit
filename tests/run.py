import unittest
from cast.application.test import run
from cast.application import create_postgres_engine
import logging
logging.root.setLevel(logging.DEBUG)


class TestIntegration(unittest.TestCase):

    def test1(self):
       
        # run(kb_name='db_apptest_local', application_name='JVTEST', engine=create_postgres_engine(user='operator',password='CastAIP',host='KSHLAP',port=2280))
        run(kb_name='db_xpdl_local', application_name='XPDL_APP', engine=create_postgres_engine(user='operator', password='CastAIP', host='KSHLAP', port=2280))   


if __name__ == "__main__":
    unittest.main()

