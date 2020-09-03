
import xlrd #import data from Excel to python
data = xlrd.open_workbook(r'C:\\Users\qic656\\Desktop\python\TM-1064.xlsx')#path
table = data.sheets()[0] 


from pyecharts.charts import Line 
import pyecharts.options as opts

attr = table.col_values(0)#x axis
v = table.col_values(2)#y axis
line= (
    Line()
    .add_xaxis(attr)#x-axis of the line
    .add_yaxis('3-stage stacked 2000-period',v, symbol='circle',symbol_size=6,
               linestyle_opts=opts.LineStyleOpts(color='red', width = 3, type_='solid'),#The line's attribute
               label_opts=opts.LabelOpts(is_show=False)#no label
               )
    .set_global_opts(
        title_opts = opts.TitleOpts(title = 'Transmission Power of SiN WG Grating', pos_left = '30%', 
                                    title_textstyle_opts=opts.TextStyleOpts(font_size=15)), #Title
        xaxis_opts = opts.AxisOpts(type_='value', name ='wavelength (um)',name_location = 'middle',
                                   name_gap= 26, position = 'top',
                                   name_textstyle_opts=opts.TextStyleOpts(font_size=14),
                                   min_ = 'dataMin', max_ ='dataMax'), #x-axis attribute
        yaxis_opts=opts.AxisOpts(type_='value', name='dB', name_location = 'middle',
                                 name_gap= 30,
                                 name_textstyle_opts=opts.TextStyleOpts(font_size=14),
                                 axistick_opts=opts.AxisTickOpts(is_show=True)),#y-axis attribute
        tooltip_opts=opts.TooltipOpts(is_show=True),
        legend_opts=opts.LegendOpts(type_=None, pos_left='right', pos_top='middle')
        )
    )
line.render(r'C:\\Users\qic656\\Desktop\python\TM-1064.html')
