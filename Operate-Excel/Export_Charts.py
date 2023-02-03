import pandas as pd
import pandas.core.frame
import xlrd
import os

from pyxlchart import Pyxlchart
from xlsxwriter.workbook import Workbook


class Export_charts(object):
    """
    根据给定数据生成图表并导出

    默认数据第一列是作为横坐标，第二列往后作为数据(标签)
    """

    def __init__(self, data, name):
        if isinstance(data, dict):
            self.data = pd.DataFrame(data)
        elif isinstance(data, pandas.core.frame.DataFrame):
            self.data = data
        else:
            print('传入数据格式错误，请传入dict或dataframe!')

        self.name = name

    def main(self):
        filepath = self.set_chart()

        xl = Pyxlchart()
        xl.WorkbookDirectory = filepath
        xl.WorkbookFilename = self.name
        xl.SheetName = 'Sheet0'
        xl.ImageFilename = "MyChart"
        xl.ExportPath = filepath
        xl.ChartName = ""
        xl.start_export()

    def _to_workbook(self):
        workbook = Workbook(self.name)  # 格式化生成一个空xlsx
        worksheet = workbook.add_worksheet('Sheet0')

        # 隐藏网格线
        worksheet.hide_gridlines(2)
        # 设置单元格格式
        xlsx_property = {
            'font_size': 9,  # 字体大小
            'bold': False,  # 是否加粗
            'align': 'center',  # 水平对齐方式
            'valign': 'vcenter',  # 垂直对齐方式
            'font_name': u'微软雅黑',
            'text_wrap': False  # 是否自动换行
        }
        cell_format = workbook.add_format(xlsx_property)

        return workbook, worksheet, cell_format

    def _generate_xlsx(self):
        data = self.data
        workbook, worksheet, cell_format = self._to_workbook()

        header = data.columns.tolist()
        for col in range(len(header)):
            worksheet.write(0, col, header[col], cell_format)

        for row in range(1, data.shape[0] + 1):
            for col in range(len(header)):
                worksheet.write(row, col, data.iloc[row - 1, col], cell_format)

        return workbook, worksheet

    def set_chart(self):
        workbook, worksheet = self._generate_xlsx()

        rows = self.data.shape[0] + 1

        column_chart = workbook.add_chart({'type': 'column'})
        column_chart.add_series({
            'name': '=Sheet0!$B$1',
            'categories': '=Sheet0!$A$2:$A$' + str(rows),
            'values': '=Sheet0!$B$2:$B$' + str(rows),
            'data_labels': {'value': True, 'font': {'name': '微软雅黑', 'size': 9}}  # 显示数据标签
        })

        # 设置图表标题
        column_chart.set_title({
            'name': 'title',
            'name_font': {'name': '微软雅黑', 'size': 12, 'blod': True},
            'overlay': False
        })
        # 设置图表示例
        column_chart.set_legend({
            'font': {'name': '微软雅黑', 'size': 9},
            'position': 'top'
        })
        # 设置横纵坐标轴
        column_chart.set_x_axis({'num_format': {'name': '微软雅黑', 'size': 9}})
        column_chart.set_y_axis({
            'name': '单位：万元',
            'name_font': {'name': '宋体', 'size': 9, 'bold': True},
            'num_format': {'name': '微软雅黑', 'size': 9},
            # 'min': 0,
            # 'max': 180,
            'line': {'none': True},
            # 'minor_unit': 4,
            # 'major_unit': 20,
            'major_gridlines': {
                'visible': True,
                'line': {'width': 0.75, 'color': '#D9D9D9'}
            },
        })
        # 设置图表大小
        column_chart.set_size({'width': 970, 'height': 576})
        # 设置图表边框，填充颜色等
        column_chart.set_chartarea({'border': {'none': True}})
        # 设置绘图区域边框，填充颜色等
        column_chart.set_plotarea({'border': {'none': True}})

        line_chart = workbook.add_chart({'type': 'line'})
        line_chart.add_series({
            'name': '=Sheet0!$E$1',
            'categories': '=Sheet0!$A$2:$A$' + str(rows),
            'values': '=Sheet0!$E$2:$E$' + str(rows),
            'line': {'width': 2.75, 'color': 'red'},
            'marker': {
                'type': 'circle',
                'size': 4,
                'border': {'color': 'red'},
                'fill': {'color': 'red'},
            },
            'y2_axis': True,
            'data_labels': {'value': True, 'font': {'name': '微软雅黑', 'size': 9}, 'position': 'above'}  # 显示数据标签
        })
        line_chart.set_y2_axis({
            'num_format': {'name': '微软雅黑', 'size': 9},
            # 'min': -0.5,
            # 'max': 3,
            'line': {'none': True},
            # 'minor_unit': 0.1,
            # 'major_unit': 0.5
        })

        column_chart.combine(line_chart)

        worksheet.insert_chart('G2', column_chart)

        workbook.close()

        filepath = os.path.abspath('.')
        return filepath
