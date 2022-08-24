"""
"""
import unittest
import calc_sla


class IntegerArithmeticTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.data_list = [
            ["2022/5/16 08:28:05", "2022/05/16 17:06:22", 518, 8.6, 7.1],
            ["2022/5/16 16:55:17", "2022/05/17 10:46:11", 1071, 17.8, 2.8],
            ["2022/5/20 16:54:42", "2022/05/23 17:15:02", 4340, 72.3, 7.8],
            ["2022/5/03 09:30:32", "2022/05/05 14:05:41", 3155, 52.6, 4.1],
            ["2022/8/23 12:05:00", "2022/8/23 21:05:00", 540, 9, 7.5],
            ["2022/8/23 07:00:00", "2022/8/24 21:00:00", 2280, 38.1, 18.5],
            ["2022/3/25 09:03:00", "2022/3/30 16:15:00", 7632, 127.2, 28],
            ["2022/3/26 11:00:00", "2022/3/29 09:17:00", 4217, 70.3, 8],
            ["2022/05/16 08:00:00", "2022/05/16 17:00:00", 540, 9.0, 7],
            ["2022/05/16 08:00:00", "2022/05/16 18:00:00", 600, 10.0, 8],
            ["2022/05/16 08:00:00", "2022/05/16 19:00:00", 660, 11.0, 9],

            # 一天按8小时算
            # ["2022/05/05 08:00:00", "2022/05/06 17:00:00", 1980, 33.0, 15],





        ]
        pass

    def testCalcSingleEvent(self):  # test method names begin with 'test'
        for i in self.data_list:
            a, b, c, d, e = i
            s = calc_sla.calc_single_event(a, b)
            print(s.total_time_delta_min, s.total_time_delta_hour, s.actual_time_delta_hour)
            assert abs(s.total_time_delta_min - c) < 0.4
            assert abs(s.total_time_delta_hour - d) < 0.4
            assert abs(s.actual_time_delta_hour - e) < 0.4


if __name__ == "__main__":
    unittest.main()
