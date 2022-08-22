"""
读excel文件
读前两列，获取开始、结束时间
是否周末，是，减去24小时
是否法定假日，是，减去24小时
是正常工作日，计算中午、下班时间是否在上班时间范围内，适当减去休息时间

示例：
创建日期	关闭日期	总解决时间（min）	总解决时间（h）	最终结果（h）
2022/5/16 08:28	2022/5/16 17:06	518	8.6	7.1
2022/5/16 16:55	2022/5/17 10:46	1071	17.9	2.9
2022/5/20 16:54	2022/5/23 17:15	4340	72.3	7.9
2022/5/3 09:30	2022/5/5 14:05	3155	52.6	4.1
"""
import datetime


class SlaModel:
    def __init__(self):
        # 创建日期: 2022/5/16 08:28
        self.start_time = ''

        # 关闭日期: 2022/5/16 17:06
        self.closed_time = ''

        # 总解决时间(min) 518
        self.total_time_delta_min = ''

        # 总解决时间(h) 8.6
        self.total_time_delta_hour = ''

        # 最终结果(h) 7.1
        self.actual_time_delta_hour = ''


def str2datetime(date_str):
    time_template = '%Y-%m-%d %H:%M:%S'
    a = datetime.datetime.strptime(date_str + ':00', time_template)
    return a


def calc_seconds(start_time, closed_time):
    delta = (closed_time - start_time).total_seconds()
    return delta


def calc_actual_time_delta(start_time, closed_time):
    pass


def calc(start, close):
    sla_model = SlaModel()

    sla_model.start_time = str2datetime(start)
    sla_model.closed_time = str2datetime(close)
    sla_model.total_time_delta_min = int(calc_seconds(sla_model.start_time, sla_model.closed_time) / 60)
    sla_model.total_time_delta_hour = int(sla_model.total_time_delta_min / 60)
    sla_model.actual_time_delta_hour = calc_actual_time_delta(sla_model.start_time, sla_model.closed_time)


def main():
    calc('', '')


if __name__ == '__main__':
    main()
