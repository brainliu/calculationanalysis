#-*-coding:utf8-*-
#user:brian
#created_at:2018/4/29 10:19
# file: dealingexcel.py
#location: china chengdu 610000
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
import xlrd

all_table_index_list=[]
#[a,b,c] 装了若干个这种list a表示sheetindex，b表示sheet行的最大计算，c表示列的最大计算

#keyword_rows,keword_cols 分别是第一列的所有行的关键字，和第一行所有列的关键字
#u"合计"   u"备注":

#一个通用的函数，求定位到你所要得到的行和列的位置，最后再把在这些求和就行了

def get_sheet_start_and_end(sh,keyword_rows,keword_cols):
    """
    :param sh:  read tables
    keyword_rows:rows keword
    keword_cols: cols keyword
    :return:  real hang and real lie
    """
    nrows = sh.nrows  # 获取行数
    ncols = sh.ncols  # 获取列数
    real_rows=-1  #真实的行
    real_cols=-1  #真实的列
    for i in range(0,nrows):          #循环第一列的每一行，出现了“合计”表示最后一行
        cell_name=sh.cell_value(i,0)
        if cell_name==keyword_rows:
            # print i
            real_rows=i
            break
    for j in range(0,ncols):   #循环第三行的每一列，表示出现了“备注”表示最后一列了
        cell_name=sh.cell_value(3,j)
        if cell_name==keword_cols:
            # print j
            real_cols=j
            break
    return real_rows,real_cols



filename="new.xls"
bk=xlrd.open_workbook(filename)
n=bk.nsheets
shrange=range(bk.nsheets)


#小计求和
def calculate_all_sum():
    reault=0.0
    print u"表格名称",u"合计",u"金额"
    for sheet in shrange:
        #sh为获取到的sheet的内容
        sh=bk.sheet_by_index(sheet)
        keyword_rows=u"合计"
        keword_cols=u"备注"
        real_rows,real_cols=get_sheet_start_and_end(sh,keyword_rows, keword_cols)
        # print  (real_rows, real_cols)
        all_table_index_list.append([sheet,real_rows,real_cols])
        #  第三行开始是抬头名称
        cell_name=sh.cell_value(real_rows, 0).decode("UTF-8").encode("utf-8")
        cell_value = sh.cell_value(real_rows, real_cols-1)
        print sh.name,cell_name,cell_value
        reault+=cell_value
    print u"所有的合计",reault
    return reault
zongheji=calculate_all_sum()
print all_table_index_list
#找出所有的材料类别有哪些呢
#材料款的类别在第三行，第一列
#所有的材料集合

all_materi_set=set()
all_finacial_category=set()
for index in all_table_index_list:
    # print index[0],"*"*20
    # sh为获取到的sheet的内容
    sh = bk.sheet_by_index(index[0])
    nrows=index[1]
    ncols = index[2]
    for i in range(3,nrows):
        temp=sh.cell_value(i, 1)
        if type(temp)==float:
            all_materi_set.add(temp)
    for j in range(1, ncols):
        temp = sh.cell_value(3, j)
        if temp not in [u"小计", u"合计"]:
            all_finacial_category.add(temp)


for category in all_finacial_category:
    print category
print all_materi_set

#所有的统计科目的集合
#采保费、材料款、运费、吊装费、延迟费、监造费、检测费

#统计每一个类别的小计，得到要查询的index，然后再求和就行

#类别的小计，显示类别的是rows-1，1
#构建一个字典
matreial_category_count={}
for id in all_materi_set:
    matreial_category_count[id]=0.0
#每一个类别的总量
sum_all=0.0
for index in all_table_index_list:
    # print index[0],"*"*20
    sh = bk.sheet_by_index(index[0])
    nrows=index[1]
    ncols = index[2]
    key = sh.cell_value(nrows-1, 1)
    value=sh.cell_value(nrows, ncols-1)
    if key in all_materi_set:
        matreial_category_count[key]+=value
        sum_all+=value

print u"所有类别合计",":",sum_all
for id in matreial_category_count:
    print u"材料类别",id,":",matreial_category_count[id]





##统计采保费、吊装费这些关键的是多少
#先统计所有分运费和吊装费等等
# 采保费5%  # 运费  # 材料款 # 修磨费  # 吊装费  # 延迟费  # 材料类别
# 采保费5.5% # 监造费  # 加工费  # 采保费 5%  # 采保费 5.5%
def get_catrgory_all(key_word):
    others=[]
    for other in all_finacial_category:
        print "ottt",other
        if other not in key_word:
            others.append(other)
    finacial_dict_couont={}
    for id in key_word:
        finacial_dict_couont[id]=0.0
    finacial_dict_couont[u"其他"]=0.0

    for index in all_table_index_list:
        # print index[0],"*"*20
        sh = bk.sheet_by_index(index[0])
        nrows=index[1]
        ncols = index[2]
        for j in range(2,ncols-1):
            key=sh.cell_value(3, j)
            value=sh.cell_value(nrows,j)
            if key in key_word:
                finacial_dict_couont[key]+=value
            else:
                if key in others:
                    finacial_dict_couont[u"其他"]+=value
    # print finacial_dict_couont
    for name in finacial_dict_couont.keys():
        print name,finacial_dict_couont[name]
    return finacial_dict_couont


#求个大类别的材料费

key_word=[u"材料款",u"采保费5%",u"采保费5.5%",u"运费",u"吊装费",u"运费",u"监造费",u"检测费",u"延迟费"]

result=get_catrgory_all(key_word)

#求材料款个大类别合计
cailiao_matreial_category_count={}
for id in all_materi_set:
    # print "leibie",id
    matreial_category_count[id]=0.0
#每一个类别的总量
cailiao_sum=0.0
x=0.0
for index in all_table_index_list:
    # print index[0],"*"*20

    sh = bk.sheet_by_index(index[0])
    nrows=index[1]
    ncols = index[2]
    key = sh.cell_value(nrows-1, 1)
    v12=0.0
    v48=0.0
    for indey in range(1,ncols):
        name=sh.cell_value(3, indey)

        if name==u"材料款":

            value = sh.cell_value(nrows, indey)

            matreial_category_count[key]+=value
            cailiao_sum+=value
            if key==12:
                print u"表名11111", sh.name, key, u"类别:", value
                v12+=value
            if key==48:
                print u"表名11111", sh.name, key, u"类别:", value
                v48+=value

print u"材料款各类别合计",":",cailiao_sum
for id in matreial_category_count:
    print u"材料类别",id,":",matreial_category_count[id]