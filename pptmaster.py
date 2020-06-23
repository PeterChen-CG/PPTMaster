# -*- coding: utf-8 -*-




import os
import time
import sys
BASE_DIR = os.path.dirname(os.path.abspath(__file__))#存放本文件所在的绝对路径
sys.path.append(BASE_DIR)
import config
import win32com.client

import pandas as pd
import numpy as np
from pandas import Series, DataFrame


# default template of pptx report
template_pptx=config.template_pptx

__all__=[
    'Report',
    'plot_table',
    'plot_chart'
]

chart_list={\
"AREA":[1,"ChartData"],\
"AREA_STACKED":[76,"ChartData"],\
"AREA_STACKED_100":[77,"ChartData"],\
"THREE_D_AREA":[-4098,"ChartData"],\
"THREE_D_AREA_STACKED":[78,"ChartData"],\
"THREE_D_AREA_STACKED_100":[79,"ChartData"],\
"BAR_CLUSTERED":[57,"ChartData"],\
"BAR_TWO_WAY":[57,"ChartData"],\
"BAR_OF_PIE":[71,"ChartData"],\
"BAR_STACKED":[58,"ChartData"],\
"BAR_STACKED_100":[59,"ChartData"],\
"THREE_D_BAR_CLUSTERED":[60,"ChartData"],\
"THREE_D_BAR_STACKED":[61,"ChartData"],\
"THREE_D_BAR_STACKED_100":[62,"ChartData"],\
"BUBBLE":[15,"BubbleChartData"],\
"BUBBLE_THREE_D_EFFECT":[87,"BubbleChartData"],\
"COLUMN_CLUSTERED":[51,"ChartData"],\
"COLUMN_STACKED":[52,"ChartData"],\
"COLUMN_STACKED_100":[53,"ChartData"],\
"THREE_D_COLUMN":[-4100,"ChartData"],\
"THREE_D_COLUMN_CLUSTERED":[54,"ChartData"],\
"THREE_D_COLUMN_STACKED":[55,"ChartData"],\
"THREE_D_COLUMN_STACKED_100":[56,"ChartData"],\
"CYLINDER_BAR_CLUSTERED":[95,"ChartData"],\
"CYLINDER_BAR_STACKED":[96,"ChartData"],\
"CYLINDER_BAR_STACKED_100":[97,"ChartData"],\
"CYLINDER_COL":[98,"ChartData"],\
"CYLINDER_COL_CLUSTERED":[92,"ChartData"],\
"CYLINDER_COL_STACKED":[93,"ChartData"],\
"CYLINDER_COL_STACKED_100":[94,"ChartData"],\
"DOUGHNUT":[-4120,"ChartData"],\
"DOUGHNUT_EXPLODED":[80,"ChartData"],\
"LINE":[4,"ChartData"],\
"LINE_MARKERS":[65,"ChartData"],\
"LINE_MARKERS_STACKED":[66,"ChartData"],\
"LINE_MARKERS_STACKED_100":[67,"ChartData"],\
"LINE_STACKED":[63,"ChartData"],\
"LINE_STACKED_100":[64,"ChartData"],\
"THREE_D_LINE":[-4101,"ChartData"],\
"PIE":[5,"ChartData"],\
"PIE_EXPLODED":[69,"ChartData"],\
"PIE_OF_PIE":[68,"ChartData"],\
"THREE_D_PIE":[-4102,"ChartData"],\
"THREE_D_PIE_EXPLODED":[70,"ChartData"],\
"PYRAMID_BAR_CLUSTERED":[109,"ChartData"],\
"PYRAMID_BAR_STACKED":[110,"ChartData"],\
"PYRAMID_BAR_STACKED_100":[111,"ChartData"],\
"PYRAMID_COL":[112,"ChartData"],\
"PYRAMID_COL_CLUSTERED":[106,"ChartData"],\
"PYRAMID_COL_STACKED":[107,"ChartData"],\
"PYRAMID_COL_STACKED_100":[108,"ChartData"],\
"RADAR":[-4151,"ChartData"],\
"RADAR_FILLED":[82,"ChartData"],\
"RADAR_MARKERS":[81,"ChartData"],\
"STOCK_HLC":[88,"ChartData"],\
"STOCK_OHLC":[89,"ChartData"],\
"STOCK_VHLC":[90,"ChartData"],\
"STOCK_VOHLC":[91,"ChartData"],\
"SURFACE":[83,"ChartData"],\
"SURFACE_TOP_VIEW":[85,"ChartData"],\
"SURFACE_TOP_VIEW_WIREFRAME":[86,"ChartData"],\
"SURFACE_WIREFRAME":[84,"ChartData"],\
"XY_SCATTER":[-4169,"XyChartData"],\
"XY_SCATTER_LINES":[74,"XyChartData"],\
"XY_SCATTER_LINES_NO_MARKERS":[75,"XyChartData"],\
"XY_SCATTER_SMOOTH":[72,"XyChartData"],\
"XY_SCATTER_SMOOTH_NO_MARKERS":[73,"XyChartData"]}

unit_dict = {"百":-2,"千":-3,"万":-4,"十万":-5,"百万":-6,"千万":-7,"亿":-8,"十亿":-9,"兆":-10}
def point_in_region(point,region):
    x=(point[0] >= region[0]) & (point[0] <= region[0]+region[2])
    y=(point[1] >= region[1]) & (point[1] <= region[1]+region[3])
    return True if (x & y) else False   

def set_default_font(text_frame,content_type = "content"):
    if content_type == "title":
        font, size, color = config.title_font, config.title_fontsize, config.title_fontcolor
    elif content_type == "summary":
        font, size, color = config.summary_font, config.summary_fontsize, config.summary_fontcolor
    elif content_type == "content":
        font, size, color = config.content_font, config.content_fontsize, config.content_fontcolor
    elif content_type == "footnote":
        font, size, color = config.footnote_font, config.footnote_fontsize, config.footnote_fontcolor    
    text_frame.TextRange.Font.Size = size
    text_frame.TextRange.Font.Name = font
    text_frame.TextRange.Font.NameFarEast = font
    text_frame.TextRange.Font.Color.RGB = color


def plot_table(slide,df,left,top,width,height,columns_names=True,index_names=True):
    '''将pandas数据框添加到slide上，并生成pptx上的表格
    输入：
    slide：PPT的一个页面，由pptx.Presentation().slides.add_slide()给定
    df：需要转换的数据框
    lef,top: 表格在slide中的位置
    width,height: 表格在slide中的大小
    index_names: Bool,是否需要显示行类别的名称
    columns_names: Bool,是否需要显示列类别的名称
    返回：
    返回table对象
    '''
    df=pd.DataFrame(df)
    rows, cols = df.shape
	
    if index_names:
        index_level = len(df.index[0]) if isinstance(df.index[0],tuple) else 1
    else:
        index_level = 0
    if columns_names:
        col_level = len(df.columns[0]) if isinstance(df.columns[0],tuple) else 1
    else:
        col_level = 0 
    #table = slide.Shapes.AddTable(rows+col_level, cols+index_level, left, top, width, height).Table
    table = slide.Shapes.AddTable(1, 1, left, top, width, height/(rows+col_level)).Table
    cell=table.Cell(1, 1).Shape.TextFrame
    cell.MarginLeft, cell.MarginRight = 0, 0
    cell.MarginTop, cell.MarginBottom = 0, 0
    cell.HorizontalAnchor = 2 # 2 as center, 1 as No alignment
    cell.VerticalAnchor = 3 # 3 as center, 4 as bottom, 1 as top			
    cell.TextRange.Font.Size = config.content_fontsize
    cell.TextRange.Font.Name = config.content_font
    cell.TextRange.Font.NameFarEast = config.content_font	

    for row in range(rows+col_level-1):
        table.Rows.Add()
    for col in range(cols+index_level-1):
        table.Columns.Add()			
			
    # 设置表格文本值
    row_parent_list, row_child_list = [rows-1],[]
    if index_level > 1:
        for i in range(index_level):
            parent_pos = 0
            for row in range(rows):
                if row == 0:
                    cell_start=table.Cell(row+col_level+1,i+1)
                    cell_start.Shape.TextFrame.TextRange.Text = '%s'% (df.index[row][i])
                    table.Cell(1,i+1).Shape.TextFrame.TextRange.Text = '%s'%(df.index.names[i])
                    table.Cell(1,i+1).Merge(table.Cell(col_level,i+1))
                    start_pos,end_pos = row, row
                else:
                    if row <= row_parent_list[parent_pos]:
                        if df.index[row][i] == df.index[row-1][i]:
                            if row < rows-1:
                                end_pos = end_pos + 1
                            else:
                                row_child_list.append(row)
                                cell_start.Merge(table.Cell(row+col_level+1,i+1))
                        else:
                            if end_pos > start_pos:
                                cell_start.Merge(table.Cell(end_pos+col_level+1,i+1))
                            row_child_list.append(end_pos)
                            cell_start=table.Cell(row+col_level+1,i+1)
                            cell_start.Shape.TextFrame.TextRange.Text = '%s'% (df.index[row][i])
                            if row < rows-1:
                                start_pos,end_pos = row, row
                            else:
                                row_child_list.append(row)
                    else:
                        if end_pos > start_pos:
                            cell_start.Merge(table.Cell(end_pos+col_level+1,i+1))
                        row_child_list.append(end_pos)
                        cell_start=table.Cell(row+col_level+1,i+1)
                        cell_start.Shape.TextFrame.TextRange.Text = '%s'% (df.index[row][i])                        
                        if row < rows-1:
                            start_pos,end_pos = row, row
                            parent_pos = parent_pos + 1
                        else:
                            row_child_list.append(row)
            row_parent_list = row_child_list
            row_child_list = []                 
  
    if index_level == 1:
        for row in range(rows):
            cell=table.Cell(row+col_level+1,1)
            cell.Shape.TextFrame.TextRange.Text = '%s'% (df.index[row])
            table.Cell(col_level,1).Shape.TextFrame.TextRange.Text = '%s'%(df.index.name)

    col_parent_list, col_child_list = [cols-1],[]
    if col_level > 1:
        for i in range(col_level):
            parent_pos = 0
            for col in range(cols):
                if col == 0:
                    cell_start=table.Cell(i+1,col+index_level+1)
                    cell_start.Shape.TextFrame.TextRange.Text = '%s'% (df.columns[col][i])
                    start_pos,end_pos = col, col
                else:
                    if col <= col_parent_list[parent_pos]:
                        if df.columns[col][i] == df.columns[col-1][i]:
                            if col < cols-1:
                                end_pos = end_pos + 1
                            else:
                                col_child_list.append(col)
                                cell_start.Merge(table.Cell(i+1,col+index_level+1))
                        else:                           
                            if end_pos > start_pos:
                                cell_start.Merge(table.Cell(i+1,end_pos+index_level+1))
                            col_child_list.append(end_pos)
                            cell_start=table.Cell(i+1,col+index_level+1)
                            cell_start.Shape.TextFrame.TextRange.Text = '%s'% (df.columns[col][i]) 
                            if col < cols-1:
                                start_pos,end_pos = col, col
                            else:
                                col_child_list.append(col)
                    else:
                        if end_pos > start_pos:
                            cell_start.Merge(table.Cell(i+1,end_pos+index_level+1))
                        col_child_list.append(end_pos)
                        cell_start=table.Cell(i+1,col+index_level+1)
                        cell_start.Shape.TextFrame.TextRange.Text = '%s'% (df.columns[col][i])
                        if col < cols-1:
                            start_pos,end_pos = col, col
                            parent_pos = parent_pos + 1
                        else:
                            col_child_list.append(col)
            col_parent_list = col_child_list
            col_child_list = [] 

    if col_level == 1:
        for col in range(cols):
            cell=table.Cell(1,col+index_level+1)
            cell.Shape.TextFrame.TextRange.Text = '%s'% (df.columns[col])

    m = df.values
    for row in range(rows):
        for col in range(cols):
            cell=table.Cell(row+col_level+1,col+index_level+1)
            if isinstance(m[row, col],float):
                cell.Shape.TextFrame.TextRange.Text = '%.2f'%(m[row, col])
            else:
                cell.Shape.TextFrame.TextRange.Text = '%s'%(m[row, col])	


    '''
    if columns_names:
        for col_index, col_name in enumerate(list(df.columns)):
            cell=table.Cell(1,col_index+index_names+1)
            cell.Shape.TextFrame.TextRange.Text = '%s'%(col_name)
    if index_names:
        for col_index, col_name in enumerate(list(df.index)):
            cell=table.Cell(col_index+columns_names+1,1)
            cell.Shape.TextFrame.TextRange.Text = '%s'%(col_name)
        table.Cell(1,1).Shape.TextFrame.TextRange.Text = '%s'%(df.index.name)

    m = df.values
    for row in range(rows):
        for col in range(cols):
            cell=table.Cell(row+columns_names+1, col+index_names+1)
            if isinstance(m[row, col],float):
                cell.Shape.TextFrame.TextRange.Text = '%.2f'%(m[row, col])
            else:
                cell.Shape.TextFrame.TextRange.Text = '%s'%(m[row, col])

    # 设置表格默认格式
    for row in range(rows+col_level):
        for col in range(cols+index_level):
            cell=table.Cell(row+1, col+1).Shape.TextFrame
            cell.MarginLeft, cell.MarginRight = 0, 0
            cell.MarginTop, cell.MarginBottom = 0, 0
            cell.HorizontalAnchor = 2 # 2 as center, 1 as No alignment
            cell.VerticalAnchor = 3 # 3 as center, 4 as bottom, 1 as top			
            cell.TextRange.Font.Size = config.content_fontsize
            cell.TextRange.Font.Name = config.content_font
            cell.TextRange.Font.NameFarEast = config.content_font		
    '''    
    return table


def plot_chart(slide,df,left,top,width,height,chart_type='COLUMN_CLUSTERED',**kwarg):
    '''根据pandas数据框制作图表添加到slide上
    输入：
    slide：PPT的一个页面，由pptx.Presentation().slides.add_slide()给定
    df：制作图表的数据
    lef,top: 图表在slide中的位置
    width,height: 图表在slide中的大小
    chart_type: 图表类型
    chart_title: 图表标题
    number_format：图表数据标签的数字格式
    返回：
    返回table对象
    '''

    chart=slide.Shapes.AddChart2(-1, chart_list[chart_type.upper()][0],left, top, width, height).Chart
    #chart.ChartData.Activate
    chart.ChartData.Workbook.Worksheets("Sheet1").Range("A1:D5").Clear()

    for i in range(chart.SeriesCollection().Count):	
        chart.SeriesCollection(1).Delete()

    df=pd.DataFrame(df)	
    rows,cols = df.shape	
    df_values = df.values.tolist()
    for i in range(len(df_values)):
        for j in range(len(df_values[i])):
            # print(type(desc_values[i][j]))
            if type(df_values[i][j]) in [np.int32,np.int64,np.float32,np.float64]:
                df_values[i][j] = np.float64(df_values[i][j])

    '''
    for row in range(rows):
        chart.ChartData.Workbook.Worksheets("Sheet1").Cells(row+2,1).Value = df.index[row]
        for col in range(cols):
            if row==0:
                chart.ChartData.Workbook.Worksheets("Sheet1").Cells(1,col+2).Value  = df.columns[col]			
            chart.ChartData.Workbook.Worksheets("Sheet1").Cells(row+2,col+2).Value = df_values[row][col]
    '''
    index_level = len(df.index[0]) if isinstance(df.index[0],tuple) else 1
    col_level = len(df.columns[0]) if isinstance(df.columns[0],tuple) else 1 
    for row in range(rows):
        if isinstance(df.index[row],tuple):
            for i in range(index_level):
                chart.ChartData.Workbook.Worksheets("Sheet1").Cells(row+col_level+1,i+1).Value = df.index[row][i]
        else:
            chart.ChartData.Workbook.Worksheets("Sheet1").Cells(row+col_level+1,1).Value = df.index[row]
        for col in range(cols):
            if row==0:
                if isinstance(df.columns[col],tuple):
                    for j in range(col_level):
                        chart.ChartData.Workbook.Worksheets("Sheet1").Cells(j+1,col+index_level+1).Value  = df.columns[col][j]			
                else:
                    chart.ChartData.Workbook.Worksheets("Sheet1").Cells(1,col+index_level+1).Value  = df.columns[col]			
            chart.ChartData.Workbook.Worksheets("Sheet1").Cells(row+col_level+1,col+index_level+1).Value = df_values[row][col]    
    region_str = "Sheet1!A1:"+chr(cols+index_level+64)+str(rows+col_level) 
    chart.SeriesCollection().Add(region_str,2,True,True)
    
	# 添加图表标题
    if "chart_title" in kwarg:
        chart.HasTitle = True
        chart.ChartTitle.Text = kwarg["chart_title"]
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = config.content_font
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.NameFarEast = config.content_font
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = config.charttitle_fontsize
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True				
    else:
        chart.HasTitle = False

    # 添加图例
    
    if df.shape[1]>1:
        chart.HasLegend = True
        chart.Legend.Format.TextFrame2.TextRange.Font.Name = config.content_font
        chart.Legend.Format.TextFrame2.TextRange.Font.NameFarEast = config.content_font
        chart.Legend.Format.TextFrame2.TextRange.Font.size = config.chartlegend_fontsize		
        chart.Legend.Position = -4160 # 2 as upper-right,-4107 as bottom,-4160 as top, -4131 as left, -4152 as right 
        chart.Legend.IncludeInLayout = False
    else:
        chart.HasLegend = False

    non_available_list=['BUBBLE','BUBBLE_THREE_D_EFFECT','XY_SCATTER',\
    'XY_SCATTER_LINES','PIE']

    # 设置X轴/Y轴
    value_axis=kwarg["value_axis"] if 'value_axis' in kwarg else True
    major_grid=kwarg["major_grid"] if 'major_grid' in kwarg else False
    unit=kwarg["unit"] if 'unit' in kwarg else False

    if (chart_type not in non_available_list):
        #chart.HasAxis[1] = True # 1 as CategoryAxis, 2 as ValueAxis, 3 as SeriesAxis
        chart.Axes(1).TickLabels.Font.Name = config.content_font
        chart.Axes(1).TickLabels.Font.Size = config.charttick_fontsize		
        if value_axis:
            #chart.HasAxis[2,1] = True
            chart.Axes(2).TickLabels.Font.Name = config.content_font
            chart.Axes(2).TickLabels.Font.Size = config.charttick_fontsize
            chart.Axes(2).MajorTickMark = 3 # 3 as outside, 2 as inside, 4 as cross, -4142 as none
            chart.Axes(2).Format.Line.ForeColor.RGB = config.RGBColor(120,120,120)
            if unit:
                chart.Axes(2).DisplayUnit = unit_dict[unit]
                chart.Axes(2).HasDisplayUnitLabel = True
                chart.Axes(2).DisplayUnitLabel.Format.TextFrame2.TextRange.Font.Name = config.content_font
                chart.Axes(2).DisplayUnitLabel.Format.TextFrame2.TextRange.Font.NameFarEast = config.content_font
                chart.Axes(2).DisplayUnitLabel.Format.TextFrame2.TextRange.Font.size = config.charttick_fontsize			
														
            if major_grid:
                chart.Axes(2).HasMajorGridlines = True
                chart.Axes(2).MajorGridlines.Border.Color.RGB = config.RGBColor(240,240,240)
            else:
                chart.Axes(2).HasMajorGridlines = False
        else:
            chart.Axes(2).Delete()
            chart.Axes(2).HasMajorGridlines = False
 
    # 添加数据标签
    number_format=kwarg["number_format"] if 'number_format' in kwarg else config.number_format_data
    if chart_type not in non_available_list:
        for i in range(chart.SeriesCollection().Count):
            chart.SeriesCollection(i+1).HasDataLabels = True
            chart.SeriesCollection(i+1).DataLabels().Format.TextFrame2.TextRange.Font.Name = config.content_font
            chart.SeriesCollection(i+1).DataLabels().Format.TextFrame2.TextRange.Font.NameFarEast = config.content_font
            chart.SeriesCollection(i+1).DataLabels().Format.TextFrame2.TextRange.Font.Size = config.chartdata_fontsize
            chart.SeriesCollection(i+1).DataLabels().NumberFormat = number_format
            if chart_type in ["LINE","LINE_MARKERS","LINE_MARKERS_STACKED","LINE_MARKERS_STACKED_100","LINE_STACKED","LINE_STACKED_100"]:
                chart.SeriesCollection(i+1).DataLabels().Position = 0 # 0 as above, 1 as below, 5 as bestfit			

    if chart_type in ["PIE","PIE_EXPLODED","PIE_OF_PIE","BAR_OF_PIE","THREE_D_PIE","THREE_D_PIE_EXPLODED"]:
        chart.SeriesCollection(1).HasDataLabels = True
        chart.SeriesCollection(1).DataLabels().Format.TextFrame2.TextRange.Font.Name = config.content_font
        chart.SeriesCollection(1).DataLabels().Format.TextFrame2.TextRange.Font.NameFarEast = config.content_font
        chart.SeriesCollection(1).DataLabels().Format.TextFrame2.TextRange.Font.Size = config.chartdata_fontsize
        chart.SeriesCollection(1).DataLabels().NumberFormat = number_format
        chart.SeriesCollection(1).DataLabels().ShowCategoryName = True
        chart.SeriesCollection(1).DataLabels().ShowPercentage = True
        #chart.SeriesCollection(1).DataLabels().Position = 0 
    
    
    return chart

class Report():
    def __init__(self,filename=None,layout_default=1,chart_type_default='COLUMN_CLUSTERED'):
        '''
        默认绘图类型后期会改为auto
        '''

        self.filename=filename
        self.layout_default=layout_default
        self.chart_type_default=chart_type_default
		#
        application = win32com.client.Dispatch("PowerPoint.Application")
        application.Visible = True
        application.DisplayAlerts = False 
    
        if filename is None:
            self.filename = "自动生成"
            if os.path.exists('template.pptx'):
                prs=application.Presentations.Open(BASE_DIR + "\\" + 'template.pptx')
            elif template_pptx is not None:
                prs=application.Presentations.Open(BASE_DIR + "\\" + template_pptx)
            else:
                prs=application.Presentations.Add()
        else :
            prs=application.Presentations.Open(BASE_DIR + "\\" + filename)
			#prs=application.Presentations.Open(BASE_DIR + "\\" + filename,False,False,False)
        self.prs=prs


    def location_suggest(self,num=1,d_type="H",content_left=None,content_top=None,content_width=None,content_height=None,location=None):
        '''统一管理slides各个模块的位置
        parameter
        --------
        num: 主体内容（如图、外链图片、文本框等）的个数，默认从左到右依次排列
        rate: 主体内容的宽度综合

        return
        -----
        locations: dict格式. l代表left,t代表top,w代表width，h代表height
        '''
        if content_left == None:
            content_left = config.content_loc[0]
        if content_top == None:
            content_top = config.content_loc[1]
        if content_width == None:
            content_width = config.content_loc[2]
        if content_height == None:
            content_height = config.content_loc[3]
        if location:
            content_left, content_width = location["l"], location["w"]
            content_top, content_height = location["t"], location["h"]

        if not(isinstance(num,list)):
            if num>1:
                if d_type == "H":
                    gap = 0.01
                    width = (content_width+gap)/num
                    lefts=[content_left+width*i for i in range(num)]
                    tops=[content_top]*num
                    widths=[width]*num
                    heights=[content_height]*num
                    locations=[{'l':lefts[i],'t':tops[i],'w':widths[i],'h':heights[i]} for i in range(num)]
                elif d_type == "V":
                    gap = 0.01
                    height = (content_height+gap)/num
                    tops=[content_top+height*i for i in range(num)]
                    lefts=[content_left]*num
                    heights=[height]*num
                    widths=[content_width]*num
                    locations=[{'l':lefts[i],'t':tops[i],'w':widths[i],'h':heights[i]} for i in range(num)]
                # 设置内容数量为多个，但是分布方式未定义时，默认为一个内容
                else:
                    locations=[{'l':content_left,'t':content_top,'w':content_width,'h':content_height}]
            else :
                locations=[{'l':content_left,'t':content_top,'w':content_width,'h':content_height}]
        else:
            n = len(num)
            if n>1:
                if d_type == "H":
                    gap = (1-sum(num))/(n-1)
                    widths=[content_width*i for i in num]
                    lefts = [content_left]
                    for i in range(n-1):
                        lefts.append(lefts[i]+widths[i]+content_width*gap) 
                    tops=[content_top]*n
                    heights=[content_height]*n
                    locations=[{'l':lefts[i],'t':tops[i],'w':widths[i],'h':heights[i]} for i in range(n)]
                elif d_type == "V":
                    gap = (1-sum(num))/(n-1)
                    heights=[content_height*i for i in num]
                    tops = [content_top]
                    for i in range(n-1):
                        tops.append(tops[i]+heights[i]+content_height*gap) 
                    lefts=[content_left]*n
                    widths=[content_width]*n
                    locations=[{'l':lefts[i],'t':tops[i],'w':widths[i],'h':heights[i]} for i in range(n)]
                else:
                    locations=[{'l':content_left,'t':content_top,'w':content_width,'h':content_height}]
            else :
                locations=[{'l':content_left,'t':content_top,'w':content_width,'h':content_height}]            
        return locations

    def add_slide(self,contents=[],title='',summary='',footnote='',contents_layout_arr=[[0,1,"H"],],locationlist=None,slide_layout_loc='auto',**kwarg):
        '''新增一个PPT页面

        contents=[{'data':,'plot_type':,'chart_type':,},] # 三个是必须字段，其他根据plot_type不同而不同
        
        contents_layout_arr=[[0,1,"H"],] #内容分布模式，[a,b,c]代表将位置列表中的第a个位置进行切分操作，可用多个列表操作多次,c="H"为横向，"V"为纵向；
        b为数字，则将第a个区域横向或纵向均分为b个区域，b为百分比组成的列表，则将第a个区域横向或纵向按百分比切分。

        locationlist = [[0.1,0.2,0.7,0.5],] #二维列表，[left,top,width,height]代表一个内容的位置信息
        
        slide_layout_loc: 母版位置，[a,b]，a为主母版索引号，b为子母版索引号
        '''
        # 选取的板式，生成页面
        if slide_layout_loc == 'auto':
            layout=self.layout_default
        slide = self.prs.Slides.AddSlide(self.prs.Slides.Count+1,self.prs.SlideMaster.CustomLayouts[layout])

        slide_width=self.prs.PageSetup.SlideWidth
        slide_height=self.prs.PageSetup.SlideHeight

        # 添加标题
        if title:        
            if slide.Shapes.Title:
                slide.Shapes.Title.TextFrame.TextRange.Text = title
            else:
                left,top = config.title_loc[0]*slide_width, config.title_loc[1]*slide_height
                width,height = config.title_loc[2]*slide_width, config.title_loc[3]*slide_height
                txBox = slide.Shapes.AddTextbox(1,left, top, width, height)
                txBox.TextFrame.TextRange.Text = title
                set_default_font(txBox.TextFrame,"title")
        # 添加结论
        if summary:
            for i in range(slide.Shapes.Count):
                if point_in_region((slide.Shapes[i].Left/slide_width,slide.Shapes[i].Top/slide_height),config.summary_region):
                    txtSummary = slide.Shapes[i]
                    txtSummary.TextFrame.TextRange.Text = summary					
                    break					

            else:
                left,top = config.summary_loc[0]*slide_width, config.summary_loc[1]*slide_height
                width,height = config.summary_loc[2]*slide_width, config.summary_loc[3]*slide_height
                txBox = slide.Shapes.AddTextbox(1,left, top, width, height)
                txBox.TextFrame.TextRange.Text = summary
                set_default_font(txBox.TextFrame,"summary")		

        # 添加脚注
        if footnote:
            for i in range(slide.Shapes.Count):
                if point_in_region((slide.Shapes[i].Left/slide_width,slide.Shapes[i].Top/slide_height),config.footnote_region):
                    txtfootnote = slide.Shapes[i]
                    txtfootnote.TextFrame.TextRange.Text = footnote					
                    break					

            else:
                left,top = config.footnote_loc[0]*slide_width, config.footnote_loc[1]*slide_height
                width,height = config.footnote_loc[2]*slide_width, config.footnote_loc[3]*slide_height
                txBox = slide.Shapes.AddTextbox(1,left, top, width, height)
                txBox.TextFrame.TextRange.Text = footnote
                set_default_font(txBox.TextFrame,"footnote")
        
        # 标准化内容部分的数据格式
        if not(isinstance(contents,list)):
            contents=[contents]
        for i,d in enumerate(contents):
            if not(isinstance(d,dict)):
                if isinstance(d,(pd.core.frame.DataFrame,pd.core.frame.Series)):
                    plot_type='chart'
                    chart_type=self.chart_type_default
                    d=pd.DataFrame(d)
                elif isinstance(d,str) and os.path.exists(d):
                    plot_type='picture'
                    chart_type=''
                elif isinstance(d,str) and not(os.path.exists(d)):
                    plot_type='textbox'
                    chart_type=''
                else:
                    print('未知的数据格式，请检查数据')
                    plot_type=''
                    chart_type=''
                contents[i]={'data':d,'plot_type':plot_type,'chart_type':chart_type}

        # 多个内容时，获取各内容的位置
        locations=[]
        if locationlist:
            for loc in range(len(locationlist)):
                locations.append({"l":loc[0],"t":loc[1],"w":loc[2],"h":loc[3]})
        else:        
            for i in range(len(contents_layout_arr)):
                if i ==0:
                    locations.extend(self.location_suggest(contents_layout_arr[0][1],contents_layout_arr[0][2]))
                else:
                    old, num, d_type = contents_layout_arr[i]
                    new=self.location_suggest(num,d_type,location=locations[old])
                    locations.pop(old)
                    for j in range(len(new)):
                        locations.insert(old+j,new[j])

        # 绘制主体部分
        for i,dd in  enumerate(contents):
            dd=dd.copy()		
            plot_type=dd.pop('plot_type')
            left,top=locations[i]['l']*slide_width,locations[i]['t']*slide_height
            width,height=locations[i]['w']*slide_width,locations[i]['h']*slide_height
            chart_type=dd.pop('chart_type') if 'chart_type' in dd else self.chart_type_default
            data=dd.pop("data")
            if plot_type in ['table']:
                # 绘制表格
                plot_table(slide,data,left,top,width,height)
            elif plot_type in ['textbox']:
                # 输出文本框
                txBox = slide.Shapes.AddTextbox(1,left, top, width, height)
                txBox.TextFrame.TextRange.Text =data
                set_default_font(txBox.TextFrame)
            elif plot_type in ['pic','picture','figure']:
                slide.Shapes.AddPicture(data, 0, 1, left, top, Height=height)
            
            elif plot_type in ['chart']:
                # 插入图表
                chart = plot_chart(slide,data,left, top, width, height, chart_type,**dd)
                chart.ChartData.Workbook.Close()



    def save(self,filename=None):
        str1 = os.path.splitext(self.filename)[0]
        filename=str1+time.strftime('_%Y%m%d%H%M.pptx', time.localtime()) if filename is None else filename
        self.prs.SaveAs(BASE_DIR + "\\" + filename)

