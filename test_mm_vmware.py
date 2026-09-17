import unittest
import io
from contextlib import redirect_stdout

import pandas as pd

from MM_VMware import clean_name, filter_and_deduplicate, parse_utilization, prepare_target


class ProcessingTests(unittest.TestCase):
    def setUp(self):
        self.enterContext(redirect_stdout(io.StringIO()))

    def test_percent_units(self):
        result = parse_utilization(pd.Series(['1%', '0,5%', '45,1%', 0.45, 1, None]))
        self.assertEqual(result.iloc[:5].tolist(), [1, 0.5, 45.1, 45, 100])
        self.assertTrue(pd.isna(result.iloc[5]))

    def test_vmware_states_without_cpu(self):
        frame = pd.DataFrame({'name': ['a', 'b', 'c'],
                              'maintenance_state': ['inMaintenance', 'notInMaintenance', 'notInMaintenance'],
                              'connection_state': ['connected', 'disconnected', 'connected']})
        self.assertEqual(filter_and_deduplicate(prepare_target(frame))['node_name'].tolist(), ['a', 'b'])

    def test_legacy_cpu_ranking(self):
        frame = pd.DataFrame({'node_name': ['a', 'a', 'a'],
                              'node_state': ['["maintenance"]'] * 3,
                              'max_cpu': [2, 3, 3], 'cpu_2sigma_usage': [9, 1, 4]})
        self.assertEqual(filter_and_deduplicate(frame).index.tolist(), [2])

    def test_empty_selection(self):
        frame = pd.DataFrame({'node_name': ['a'], 'node_state': ['connected']})
        self.assertTrue(filter_and_deduplicate(frame).empty)

    def test_ip_is_not_shortened(self):
        self.assertEqual(clean_name(' 10.20.30.40 '), '10.20.30.40')
        self.assertEqual(clean_name('HOST.example.org'), 'host')


if __name__ == '__main__':
    unittest.main()
