from pyoffice import PPT
from wsdata.consts import ZHUZHAI, BIESHU, SHANGYE, BANGONG
from wsdata.models import WinsunDatabase
from wsdata.utils import Week


class Report:
    ws = WinsunDatabase()
    ppt = PPT('template.pptx')
    week = Week().N

    def trend(self, usage):
        """返回指定物业类型的走势数据与说理"""

        def index_adjust(label):
            return '-'.join(f'{x[4:6]}.{x[-2:]}' for x in label.split('-'))

        gxj = self.ws.gxj('trend', 'week', 10, usage=usage)
        shuoli = gxj.shuoli(0)

        # 走势数据
        df = gxj.df_adjusted
        df.index = [index_adjust(x) for x in df.index]
        df.columns = ['上市(万㎡)', '成交(万㎡)', '均价(元/㎡)']

        # 说理
        text = '本周' + shuoli.full_text.replace('。', '；', 2)

        return text, df

    def plate_df(self, usage):
        """返回指定物业类型的分版块数据"""
        gxj = self.ws.gxj('plate', 'week', 1, usage=usage)
        df = gxj.df_adjusted
        df.columns = ['上市(万㎡)', '成交(万㎡)', '均价(元/㎡)']
        return df

    def rank(self, usage, type_):
        """ 排行榜
        :param usage: str or list 物业类型
        :param type_: str "sale":上市; "sold":成交
        :return df: 表格内容数据
        """
        # 参数
        group_by = ['板块', 'popularizename']
        outputs = ['面积', '件数']
        columns = ['排名', '板块', '项目', '面积(㎡)', '套数']

        # 成交排行添加均价
        if type_ == 'sold':
            outputs.append('均价')
            columns.append('均价(元/㎡)')

        # 查询与调整
        df = self.ws.rank(f'week_{type_}', 1, group_by, outputs, usage=usage).head(3)
        df.面积 = df.面积.round().astype(int)
        if type_ == 'sold':
            df.均价 = df.均价.astype(int)
        df.columns = columns

        # 非住宅类的再加上类型列
        if not usage == ZHUZHAI:
            columns.insert(3, '类型')
            df = df.reindex(columns=columns)

            # 每个项目的各个物业类型用"/"串联，并调整名称
            def join_usages(name):
                filter_ = (df_.popularizename == name)
                text = '/'.join(df_[filter_].功能)
                return text.replace('公寓办公', '公寓').replace('别墅', '')

            # 查询类型并整理
            df_ = self.ws.rank(f'week_{type_}', 1, ['板块', 'popularizename', '功能'], '面积', usage=usage)
            usages = [join_usages(name) for name in df.项目]
            df.类型 = usages

        return df

    def one_page(self, page_idx, usage):
        usage_label = {
            ZHUZHAI: '住宅',
            BIESHU: '别墅',
            SHANGYE: '商业',
            BANGONG: '办公'
        }[usage]

        # 获得数据
        text, df_trend = self.trend(usage)
        df_plate = self.plate_df(usage)
        rank_sale = self.rank(usage, 'sale')
        rank_sold = self.rank(usage, 'sold')

        # 填入ppt
        for shape_idx, value in [
            # 结论
            (2, text),
            # 走势图
            (4, df_trend),
            # 分板块图
            (5, f'2018年第{self.week}周南京{usage_label}市场分板块供销量价'),
            (6, df_plate),
            # 上市排行
            (7, f'2018年第{self.week}周{usage_label}市场上市面积前三'),
            (8, rank_sale),
            # 成交排行
            (9, f'2018年第{self.week}周{usage_label}市场成交面积前三'),
            (10, rank_sold)
        ]:
            self.ppt[f'{page_idx} {shape_idx}'] = value

        print(f'[*] {usage}页面完成！')


if __name__ == '__main__':
    report = Report()
    for page_idx, usage in enumerate([ZHUZHAI, BIESHU, SHANGYE, BANGONG]):
        report.one_page(page_idx, usage)
    report.ppt.save(f'E:/周报测试/2018年第{report.week}周周报.pptx')
