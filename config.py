# -*- coding: utf-8 -*-

import os
import sys
#from pptx.util import Inches, Pt
#from pptx.dml.color import RGBColor

_thisdir = os.path.realpath(os.path.split(__file__)[0])

def RGBColor(r,g,b):
    colorInt = r + (g * 256) + (b * 256 * 256)
    return colorInt

def _get_element_path(dir_name,suffix=None):
    # 目录不存在
    if not(os.path.exists(os.path.join(_thisdir,dir_name))):
        element_path=None
        return element_path

    element_path=None
    filelist=os.listdir(os.path.join(_thisdir,dir_name))
    if isinstance(suffix,str):
        suffix=[suffix]
    elif suffix is not None:
        suffix=list(suffix)
    # 寻找目录下符合后缀要求的文件，多个将被最后一个覆盖
    for f in filelist:
        if isinstance(suffix,list) and os.path.splitext(f)[1][1:] in suffix:
            element_path=os.path.join(_thisdir,dir_name,f)
    return element_path

# 使用template目录中的PPTX文件作为默认模板
template_pptx=_get_element_path('template',suffix=['pptx'])

# 使用font目录中的字体文件作为默认字体，若没有，则使用下列系统字体
font_path=_get_element_path('font',suffix=['ttf','ttc'])
if font_path is None:
    if sys.platform.startswith('win'):
        # 若能找到下列字体，优先使用后面的。font_path='C:\\windows\\fonts\\msyh.ttc'
        fontlist=['calibri.ttf','simfang.ttf','simkai.ttf','simhei.ttf','simsun.ttc','msyh.ttf','MSYH.TTC','msyh.ttc']
        for f in fontlist:
            if os.path.exists(os.path.join('C:\\windows\\fonts\\',f)):
                font_path=os.path.join('C:\\windows\\fonts\\',f)

#  默认字体大小
title_fontsize, summary_fontsize, content_fontsize, footnote_fontsize = [14,12,10,9]
charttitle_fontsize, chartdata_fontsize, charttick_fontsize, chartlegend_fontsize = [12,10,9,9]
title_fontcolor, summary_fontcolor, content_fontcolor, footnote_fontcolor= [
    RGBColor(0,0,0),RGBColor(0,0,0),RGBColor(0,0,0),RGBColor(127,127,127)
]
title_font, summary_font, content_font, footnote_font= ["微软雅黑"]*4

# PPT图表中的数字位数
chart_number_format = {
    "#,##0": ["销量","数量"]
}
number_format_data='0'

# PPT图表中坐标轴的数字标签格式
number_format_tick='0'

#  PPT中标题的默认位置，四个值依次为left、top、width、height
title_loc=[0.01,0.01,0.70,0.1]

#  PPT中结论的默认位置，四个值依次为left、top、width、height
summary_loc=[0.025,0.1,0.95,0.15]
summary_region=[0,0.08,0.6,0.1]
#  PPT中正文内容的默认位置，四个值依次为left、top、width、height
content_loc=[0.025,0.25,0.95,0.60]

#  PPT中脚注的默认位置，四个值依次为left、top、width、height
footnote_loc=[0.025,0.95,0.80,0.05]
footnote_region=[0,0.9,0.6,0.1]
