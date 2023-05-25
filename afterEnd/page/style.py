# --------------样式---------------- #
import xlwt
class style(object):
    def __init__(self):
        pass

    def style0(self):
        '''
               创建单元格格式：（默认垂直居中）
               :param font_name: 字体名称
               :param font_height: 字体高度
               :param bold: 默认加粗
               :param horz: 水平对齐方式，默认水平居中：2，左对齐：1，右对齐：3
               :return: 返回设置好的格式
               '''
        # 初始化样式
        style0 = xlwt.XFStyle()
        font = style0.font
        font.name = "微软雅黑"
        font.height = 300  # 字体大小
        alignment = style0.alignment
        alignment.horz = 2
        alignment.vert = 1
        borders=style0.borders
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1

        return style0

    def style1(self):

        style1 = xlwt.XFStyle()
        # 首行加粗、字体放大
        font = style1.font  # 11为字号，20为衡量单位
        font.name = "微软雅黑"
        font.height = 300
        font.bold = True
        alignment = style1.alignment
        alignment.horz =2
        alignment.vert = 1
        borders = style1.borders
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1


        return style1

    def style2(self):
        style2 = xlwt.XFStyle()
        # 首行加粗、字体放大
        font = style2.font
        font.name = "微软雅黑"
        font.height = 300
        font.bold = True
        alignment = style2.alignment
        alignment.horz = 2
        alignment.vert = 1
        borders = style2.borders
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        pattern = style2.pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 5

        return style2

    def style3(self):
        style3 = xlwt.XFStyle()
        # 首行加粗、字体放大
        font = style3.font
        font.name = "微软雅黑"
        font.height = 300
        font.bold = True
        alignment = style3.alignment
        alignment.horz = 2
        alignment.vert = 1
        borders = style3.borders
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        pattern=style3.pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 2

        return style3


