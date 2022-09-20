from pyecharts import options as opts
from pyecharts.charts import Map, Timeline
from pyecharts.charts import Bar, Line, Pie, Page
from pyecharts.charts import WordCloud
from pyecharts.globals import SymbolType
from pyecharts.commons.utils import JsCode
import pandas as pd

# 省份
province_lie = ['安徽的新增', '北京的新增', '重庆的新增', '福建的新增', '甘肃的新增', '广东的新增', '广西的新增', '贵州的新增', '海南的新增', '河北的新增', '河南的新增', '黑龙江的新增', '湖北的新增', '湖南的新增', '江西的新增', '吉林的新增', '江苏的新增', '辽宁的新增', '内蒙古的新增', '宁夏的新增', '青海的新增', '山西的新增', '山东的新增', '陕西的新增', '上海的新增', '四川的新增', '天津的新增', '西藏的新增', '新疆的新增', '云南的新增', '浙江的新增', '香港新增病例', '澳门新增病例', '台湾新增病例']
province = ['安徽', '北京', '重庆', '福建', '甘肃', '广东', '广西', '贵州', '海南', '河北', '河南', '黑龙江', '湖北', '湖南', '江西', '吉林', '江苏', '辽宁', '内蒙古', '宁夏', '青海', '山西', '山东', '陕西', '上海', '四川', '天津', '西藏', '新疆', '云南', '浙江', '香港', '澳门', '台湾']
province_land = ['安徽的新增', '北京的新增', '重庆的新增', '福建的新增', '甘肃的新增', '广东的新增', '广西的新增', '贵州的新增', '海南的新增', '河北的新增', '河南的新增', '黑龙江的新增', '湖北的新增', '湖南的新增', '江西的新增', '吉林的新增', '江苏的新增', '辽宁的新增', '内蒙古的新增', '宁夏的新增', '青海的新增', '山西的新增', '山东的新增', '陕西的新增', '上海的新增', '四川的新增', '天津的新增', '西藏的新增', '新疆的新增', '云南的新增', '浙江的新增']
province2 = ['安徽', '北京', '重庆', '福建', '甘肃', '广东', '广西', '贵州', '海南', '河北', '河南', '黑龙江', '湖北', '湖南', '江西', '吉林', '江苏', '辽宁', '内蒙古', '宁夏', '青海', '山西', '山东', '陕西', '上海', '四川', '天津', '西藏', '新疆', '云南', '浙江']
# 设置列对齐
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
# 打开文件
df = pd.read_excel('personal.xlsx', sheet_name='Sheet')
data2 = df['日期']
data2_list = list(data2)
data3 = df['本土新增']
data3_list = list(data3)
xList = ['香港新增病例']
xList_list = list(xList)
aList = ['澳门新增病例']
aList_list = list(aList)
tList = ['台湾新增病例']
tList_list = list(tList)
Anhui = ['安徽的新增']
Anhui_list = list(Anhui)
Beijing = ['北京的新增']
Beijing_list = list(Beijing)
Chongqing = ['重庆的新增']
Chongqing_list = list(Chongqing)
Fujian = ['福建的新增']
Fujian_list = list(Fujian)
Gansu = ['甘肃的新增']
Gansu_list = list(Gansu)
Guangdong = ['广东的新增']
Guangdong_list = list(Guangdong)
Guangxi = ['广西的新增']
Guangxi_list = list(Guangxi)
Guizhou = ['贵州的新增']
Guizhou_list = list(Guizhou)
Hainan = ['海南的新增']
Hainan_list = list(Hainan)
Hebei = ['河北的新增']
Hebei_list = list(Hebei)
Henan = ['河南的新增']
Henan_list = list(Henan)
Heilongjiang = ['黑龙江的新增']
Heilongjiang_list = list(Heilongjiang)
Hubei = ['湖北的新增']
Hubei_list = list(Hubei)
Hunan = ['湖南的新增']
Hunan_list = list(Hunan)
Jilin = ['吉林的新增']
Jilin_list = list(Jilin)
Jiangsu = ['江苏的新增']
Jiangsu_list = list(Jiangsu)
Jiangxi = ['江西的新增']
Jiangxi_list = list(Jiangxi)
Liaoning = ['辽宁的新增']
Liaoning_list = list(Liaoning)
Neimenggu = ['内蒙古的新增']
Neimenggu_list = list(Neimenggu)
Ningxia = ['宁夏的新增']
Ningxia_list = list(Ningxia)
Qinghai = ['青海的新增']
Qinghai_list = list(Qinghai)
Shandong = ['山东的新增']
Shandong_list = list(Shandong)
Shanxi = ['山西的新增']
Shanxi_list = list(Shanxi)
Shanxii = ['陕西的新增']
Shanxii_list = list(Shanxii)
Shanghai = ['上海的新增']
Shanghai_list = list(Shanghai)
Sichuan = ['四川的新增']
Sichuan_list = list(Sichuan)
Tianjin = ['天津的新增']
Tianjin_list = list(Tianjin)
Xizang = ['西藏的新增']
Xizang_list = list(Xizang)
Xinjiang = ['新疆的新增']
Xinjiang_list = list(Xinjiang)
Yunnan = ['云南的新增']
Yunnan_list = list(Yunnan)
Zhejiang = ['浙江的新增']
Zhejiang_list = list(Zhejiang)
dn = pd.read_excel('personal.xlsx', sheet_name='Sheet2')
# 对省份进行统计

data4 = dn['本土无症状']
data4_list = list(data4)

Anhui2 = ['安徽无症状']
Anhui2_list = list(Anhui2)
Beijing2 = ['北京无症状']
Beijing2_list = list(Beijing2)
Chongqing2 = ['重庆无症状']
Chongqing2_list = list(Chongqing2)
Fujian2 = ['福建无症状']
Fujian2_list = list(Fujian2)
Gansu2 = ['甘肃无症状']
Gansu2_list = list(Gansu2)
Guangdong2 = ['广东无症状']
Guangdong2_list = list(Guangdong2)
Guangxi2 = ['广西无症状']
Guangxi2_list = list(Guangxi2)
Guizhou2 = ['贵州无症状']
Guizhou2_list = list(Guizhou2)
Hainan2 = ['海南无症状']
Hainan2_list = list(Hainan2)
Hebei2 = ['河北无症状']
Hebei2_list = list(Hebei2)
Henan2 = ['河南无症状']
Henan2_list = list(Henan2)
Heilongjiang2 = ['黑龙江无症状']
Heilongjiang2_list = list(Heilongjiang2)
Hubei2 = ['湖北无症状']
Hubei2_list = list(Hubei2)
Hunan2 = ['湖南无症状']
Hunan2_list = list(Hunan2)
Jilin2 = ['吉林无症状']
Jilin2_list = list(Jilin2)
Jiangsu2 = ['江苏无症状']
Jiangsu2_list = list(Jiangsu2)
Jiangxi2 = ['江西无症状']
Jiangxi2_list = list(Jiangxi2)
Liaoning2 = ['辽宁无症状']
Liaoning2_list = list(Liaoning2)
Neimenggu2 = ['内蒙古无症状']
Neimenggu2_list = list(Neimenggu2)
Ningxia2 = ['宁夏无症状']
Ningxia2_list = list(Ningxia2)
Qinghai2 = ['青海无症状']
Qinghai2_list = list(Qinghai2)
Shandong2 = ['山东无症状']
Shandong2_list = list(Shandong2)
Shanxi2 = ['山西无症状']
Shanxi2_list = list(Shanxi2)
Shanxii2 = ['陕西无症状']
Shanxii2_list = list(Shanxii2)
Shanghai2 = ['上海无症状']
Shanghai2_list = list(Shanghai2)
Sichuan2 = ['四川无症状']
Sichuan2_list = list(Sichuan2)
Tianjin2 = ['天津无症状']
Tianjin2_list = list(Tianjin2)
Xizang2 = ['西藏无症状']
Xizang2_list = list(Xizang2)
Xinjiang2 = ['新疆无症状']
Xinjiang2_list = list(Xinjiang2)
Yunnan2 = ['云南无症状']
Yunnan2_list = list(Yunnan2)
Zhejiang2 = ['浙江无症状']
Zhejiang2_list = list(Zhejiang2)

prov_list = [xList_list, aList_list, tList_list, Anhui_list, Beijing_list, Chongqing_list, Fujian_list, Gansu_list, Guangdong_list, Guangxi_list, Guizhou_list, Hainan_list, Hebei_list, Henan_list, Heilongjiang_list, Hubei_list, Hunan_list, Jilin_list, Jiangsu_list, Jiangxi_list, Liaoning_list, Neimenggu_list, Ningxia_list, Qinghai_list, Shandong_list, Shanxi_list, Shanxii_list, Shanghai_list, Sichuan_list, Tianjin_list, Xizang_list, Xinjiang_list, Yunnan_list, Zhejiang_list]
prov_list2 = [Anhui2_list, Beijing2_list, Chongqing2_list, Fujian2_list, Gansu2_list, Guangdong2_list, Guangxi2_list, Guizhou2_list, Hainan2_list, Hebei2_list, Henan2_list, Heilongjiang2_list, Hubei2_list, Hunan2_list, Jilin2_list, Jiangsu2_list, Jiangxi2_list, Liaoning2_list, Neimenggu2_list, Ningxia2_list, Qinghai2_list, Shandong2_list, Shanxi2_list, Shanxii2_list, Shanghai2_list, Sichuan2_list, Tianjin2_list, Xizang2_list, Xinjiang2_list, Yunnan2_list, Zhejiang2_list]

day_newcreate_sym = []
day_newcreate_land = []
day_hot_land = []

#新增病例地图
tl = Timeline()
tl.add_schema(play_interval=1000, label_opts=opts.series_options.LabelOpts(is_show=True, color='white', font_size=14))
for month in range(8, 10):
    for day in range(1, 32):
        try:
            timez = df.iloc[:, :][df.日期 == (str(month) + '月' + str(day) + '日')]
            day_date = []
            for t in province_lie:
                day_date.append(list(timez[t]))
            day_newcreate_sym = []
            for t in day_date:
                x = t
                day_newcreate_sym.append(x[0])
            # print(day_newcreate_sym)
        except:
            print('没有该日期的{}.{}'.format(month, day))
        try:
            map0 = (
                Map(init_opts=opts.InitOpts(width="500px", height="300px", theme='light'))
                    .add("新增病例", [list(z) for z in zip(province, day_newcreate_sym)], "china")
                    .set_series_opts(label_opts=opts.LabelOpts(is_show=False),
                                     itemstyle_opts=opts.ItemStyleOpts(border_color='#46BEE9', opacity=0.7),
                                     )
                    .set_global_opts(
                        #title_opts=opts.TitleOpts(title="Map-{}年某些数据".format(i)),
                        legend_opts=opts.LegendOpts(textstyle_opts=opts.TextStyleOpts(color='white'), pos_top='20%'),
                        visualmap_opts=opts.VisualMapOpts(
                            pos_left='10%',
                            pos_bottom='3%',
                            range_color=['#75A9E1', '#1B3DA7', '#183895']
                                                      )
                    )
            )
            tl.add(map0, "{}月{}日".format(month, day))
        except:
            pass
tl.render('tl.html')

#新增病例和无症状柱状图
c = (
        Bar(init_opts=opts.InitOpts(width="480px", height="200px", theme='light'))
        .add_xaxis(data2_list)
        .add_yaxis("本土无症状", data4_list, itemstyle_opts={
                    "normal": {
                        "color": JsCode(
                            """new echarts.graphic.LinearGradient(0, 0, 0, 1, [{
                        offset: 0,
                        color: 'rgba(0, 244, 255, 1)'
                    }, {
                        offset: 1,
                        color: 'rgba(0, 77, 167, 1)'
                    }], false)"""
                        ),
                        "barBorderRadius": [30, 30, 30, 30],
                        "shadowColor": "rgb(0, 160, 221)",
                    }
                })
        .add_yaxis("本土新增病例", data3_list, itemstyle_opts={
                    "normal": {
                        "color": JsCode(
                            """new echarts.graphic.LinearGradient(0, 0, 0, 1, [{
                        offset: 0,
                        color: '#A1C5EA'
                    }, {
                        offset: 1,
                        color: '#F1F7FD'
                    }], false)"""
                        ),
                        "barBorderRadius": [30, 30, 30, 30],
                        "shadowColor": "rgb(0, 160, 221)",
                    }
                })
        .set_series_opts(label_opts=opts.LabelOpts(color='white'))
        .set_global_opts(
            #title_opts=opts.TitleOpts(title="Bar-DataZoom（slider-水平）"),
            datazoom_opts=[opts.DataZoomOpts(range_start=0, range_end=25)],
            legend_opts=opts.LegendOpts(textstyle_opts=opts.TextStyleOpts(color='white')),
            xaxis_opts=opts.AxisOpts(
                type_="category",
                axispointer_opts=opts.AxisPointerOpts(is_show=True, type_="shadow"),
                axislabel_opts=opts.LabelOpts(formatter=" {value}", color="#999999")
            ),
            yaxis_opts=opts.AxisOpts(
                min_='dataMin',  # y轴最小值
                axislabel_opts=opts.LabelOpts(formatter=" {value}人", color="#999999")
            ),  # y轴符号

        )
    )
c.render('c.html')

#新增病例和无症状折线图
xian = (
    Line(init_opts=opts.InitOpts(width="700px", height="400px", theme='walden'))
    .add_xaxis(data2_list)
    .add_yaxis("本土无症状", data4_list, is_smooth=True, linestyle_opts=opts.LineStyleOpts(width=5),)
    .add_yaxis("本土新增病例", data3_list, is_smooth=True, linestyle_opts=opts.LineStyleOpts(width=5))
    .set_global_opts(
        title_opts=opts.TitleOpts(title=""),
        tooltip_opts=opts.TooltipOpts(
            is_show=True,
            trigger="axis",
            axis_pointer_type="cross"
        ),
        legend_opts=opts.LegendOpts(textstyle_opts=opts.TextStyleOpts(color='white')),
        xaxis_opts=opts.AxisOpts(
            type_="category",
            axispointer_opts=opts.AxisPointerOpts(is_show=True, type_="shadow")
        ),
        yaxis_opts=opts.AxisOpts(
            min_='dataMin', #y轴最小值
            axislabel_opts=opts.LabelOpts(formatter=" {value}人")
        ),#y轴符号
        datazoom_opts=[opts.DataZoomOpts(range_start=0, range_end=25)],
    )
)


#热点词云
try:
    time15 = df.iloc[0:1, :]
    #print(time15)
    time14 = df.iloc[1:2, :]
    #print(time14)
    day_day = []

    for t in province_land:
        day_day.append(list(time15[t])[0]-list(time14[t])[0])
    print(day_day)
    day_hot_land = []
    for t in day_day:
        x = t
        if x < 0:
            x = 0
        day_hot_land.append(x)
    # print(day_newcreate_sym)
except:
    print()

cloud = (
    WordCloud(init_opts=opts.InitOpts(width="700px", height="400px"))
    .add("", [list(z) for z in zip(province2, day_hot_land)], word_size_range=[10, 50], shape=SymbolType.DIAMOND, textstyle_opts=opts.TextStyleOpts(color='white'))
    .set_series_opts(
        itemstyle_opts=opts.ItemStyleOpts(border_color='#46BEE9', opacity=0.7, color='white'),
        textstyle_opts=opts.TextStyleOpts(color='white', align='right'))
    .set_global_opts(
        title_opts=opts.TitleOpts(title="今日热点(相对于昨日的新增)", pos_left='center', title_textstyle_opts=opts.TextStyleOpts(color='white')),

    )
)
cloud.render('cloud.html')


#当日各省新增病例饼状图
try:
    timex = df.iloc[:, :][df.日期 == ("9月17日")]
    day = []
    for t in province_land:
        day.append(list(timex[t]))
    day_newcreate_land = []
    for t in day:
        x = t
        day_newcreate_land.append(x[0])
    # print(day_newcreate_sym)
except:
    print()

bing = (
    Pie(init_opts=opts.InitOpts(width="700px", height="400px", theme='light'))
    .add(
        "",
        [
            list(z)
            for z in zip(
                province2, day_newcreate_land
            )
        ],
        center=["40%", "50%"],
    )
    .set_global_opts(
        title_opts=opts.TitleOpts(title="今日本土新增", title_textstyle_opts=opts.TextStyleOpts(color='white')),
        legend_opts=opts.LegendOpts(textstyle_opts=opts.TextStyleOpts(color='white'), type_="scroll", pos_left="80%", orient="vertical"),

    )
    .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c}"))

)

#背景图
bg1 = (
    Line(init_opts=opts.InitOpts(width="1250px",
                                 height="700px",
                                 bg_color={"type": "pattern", "image": JsCode("img"), "repeat": "no-repeat"}))
    .add_xaxis([None])
    .add_yaxis("", [None])
    .set_global_opts(

        yaxis_opts=opts.AxisOpts(is_show=False),
        xaxis_opts=opts.AxisOpts(is_show=False))
)
bg1.add_js_funcs(
    """
       var img = new Image(); 
       img.src = './image/head.png';
       img.style.width = "100px";
       img.style.height = "100px";
    """
)

bg2 = (
    Line(init_opts=opts.InitOpts(width="1250px",
                                 height="700px",
                                 bg_color={"type": "pattern", "image": JsCode("img"), "repeat": "no-repeat"}))
    .add_xaxis([None])
    .add_yaxis("", [None])
    .set_global_opts(

        yaxis_opts=opts.AxisOpts(is_show=False),
        xaxis_opts=opts.AxisOpts(is_show=False))
)
bg2.add_js_funcs(
    """
       var img = new Image(); 
       img.src = './image/border.png';
       img.style.width = "100px";
       img.style.height = "100px";
    """
)

bg3 = (
    Line(init_opts=opts.InitOpts(width="1250px",
                                 height="700px",
                                 bg_color={"type": "pattern", "image": JsCode("img"), "repeat": "no-repeat"}))
    .add_xaxis([None])
    .add_yaxis("", [None])
    .set_global_opts(

        yaxis_opts=opts.AxisOpts(is_show=False),
        xaxis_opts=opts.AxisOpts(is_show=False))
)
bg3.add_js_funcs(
    """
       var img = new Image(); 
       img.src = './image/mapbg.png';
       img.style.width = "100px";
       img.style.height = "100px";
    """
)

bg4 = (
    Line(init_opts=opts.InitOpts(width="1250px",
                                 height="700px",
                                 bg_color={"type": "pattern", "image": JsCode("img"), "repeat": "no-repeat"}))
    .add_xaxis([None])
    .add_yaxis("", [None])
    .set_global_opts(

        yaxis_opts=opts.AxisOpts(is_show=False),
        xaxis_opts=opts.AxisOpts(is_show=False))
)
bg4.add_js_funcs(
    """
       var img = new Image(); 
       img.src = './image/border.png';
       img.style.width = "100px";
       img.style.height = "100px";
    """
)

bg5 = (
    Line(init_opts=opts.InitOpts(width="1250px",
                                 height="700px",
                                 bg_color={"type": "pattern", "image": JsCode("img"), "repeat": "no-repeat"}))
    .add_xaxis([None])
    .add_yaxis("", [None])
    .set_global_opts(

        yaxis_opts=opts.AxisOpts(is_show=False),
        xaxis_opts=opts.AxisOpts(is_show=False))
)
bg5.add_js_funcs(
    """
       var img = new Image(); 
       img.src = './image/border.png';
       img.style.width = "100px";
       img.style.height = "100px";
    """
)

bg6 = (
    Line(init_opts=opts.InitOpts(width="1250px",
                                 height="700px",
                                 bg_color={"type": "pattern", "image": JsCode("img"), "repeat": "no-repeat"}))
    .add_xaxis([None])
    .add_yaxis("", [None])
    .set_global_opts(

        yaxis_opts=opts.AxisOpts(is_show=False),
        xaxis_opts=opts.AxisOpts(is_show=False))
)
bg6.add_js_funcs(
    """
       var img = new Image(); 
       img.src = './image/border.png';
       img.style.width = "100px";
       img.style.height = "100px";
    """
)

line3 = (
    Line(init_opts=opts.InitOpts(width="1250px",
                                 height="700px",
                                 bg_color={"type": "pattern", "image": JsCode("img")}))
    .add_xaxis([None])
    .add_yaxis("", [None])
    .set_global_opts(
        title_opts=opts.TitleOpts(title="2022疫情数据分析",
                                  subtitle='更新日期: 2022/9/17',
                                  pos_left='center',
                                  title_textstyle_opts=opts.TextStyleOpts(font_size=25, color='white'),
                                  pos_top='5%'),
        yaxis_opts=opts.AxisOpts(is_show=False),
        xaxis_opts=opts.AxisOpts(is_show=False))
)
line3.add_js_funcs(
    """
       var img = new Image(); 
       img.src = './image/bg_4.jpg';
       
    """
)

page = Page(layout=Page.DraggablePageLayout, page_title="疫情可视化")

# 在页面中添加图表
page.add(
    line3,
    bg1,
    bg4,
    bg5,
    bg6,
    bg2,
    bg3,
    tl,
    c,
    bing,
    xian,
    cloud, )

#page.render('test.html')
# 重新布局
Page.save_resize_html("test.html",
                      cfg_file="./chart_config (7).json",
                      dest="test.html_re2.html")