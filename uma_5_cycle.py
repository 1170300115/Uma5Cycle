import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import itertools
import numpy as np

file_name=r'./relation.xlsx'
output_file_name=r'./result.xlsx'
exist_uma_sheet_name='适应性相性表'
parent_sheet_suffix='父母相性'
grand_parent_sheet_suffix='祖辈相性'

def read_index(sheet):
    row_index=dict()
    for index,value in sheet.iloc[:,0].items():
        row_index[value]=index
    return row_index

def put_data(uma_name):
    return uma_name,Uma(uma_name)

class Uma(object):
    
    def __init__(self,name):
        self.name=name
        self.parent_sheet = pd.read_excel(file_name,name+parent_sheet_suffix)
        self.grand_parent_sheet = pd.read_excel(file_name,sheet_name=name+grand_parent_sheet_suffix)
        self.parent_index = read_index(self.parent_sheet)
        self.grand_parent_index=read_index(self.grand_parent_sheet)
    
    def get_parent_score(self,parent_a,parent_b):
        column=self.parent_sheet[parent_a.name]
        return int(column[self.parent_index[parent_b.name]])
    
    def get_grand_parent_score(self,parent,grand_parent):
        column=self.grand_parent_sheet[parent.name]
        return int(column[self.grand_parent_index[grand_parent.name]])
    
class ScoreData(object):
    def __init__(self,uma_list):
        self.uma_length = len(uma_list)
        self.uma_list = uma_list
    def getDataList(self):
        result =list()
        data = list()
        range_sort=list(itertools.permutations(range(1,self.uma_length),self.uma_length-1))
        for item in range_sort:
            index=[0]+list(item)
            total_score=0
            score_list=list()
            name_list=list()
            for i in range(0,self.uma_length):
                uma=self.uma_list[index[i]]
                parent_1=self.uma_list[index[i-1]]
                parent_2=self.uma_list[index[i-2]]
                grand_parent_1_1=self.uma_list[index[i-2]]
                grand_parent_1_2=self.uma_list[index[i-3]]
                grand_parent_2_1=self.uma_list[index[i-3]]
                grand_parent_2_2=self.uma_list[index[i-4]]
                score=uma.get_parent_score(parent_1,parent_2)
                score+=uma.get_grand_parent_score(parent_1,grand_parent_1_1)
                score+=uma.get_grand_parent_score(parent_1,grand_parent_1_2)
                score+=uma.get_grand_parent_score(parent_2,grand_parent_2_1)
                score+=uma.get_grand_parent_score(parent_2,grand_parent_2_2)
                score_list.append(score)
                name_list.append(uma.name)
                total_score+=score
            if(len(data)==0 or data[-1]<total_score):
                data=name_list+score_list+[total_score]
                result=list()
                result.append(data)
            elif data[-1] == total_score:
                result.append(name_list+score_list+[total_score])
        return result
    
            
class Calculator(object):

    def __init__(self,option,must_contain=set()):
        sheet =pd.read_excel(file_name,exist_uma_sheet_name)
        index=read_index(sheet)  
        self.exist_uma=dict()
        self.must_contain=must_contain
        with ThreadPoolExecutor(max_workers=10) as executor:
            print('开始读取马娘相性数据')
            future_task_list=list()
            for i in range(1,sheet.columns.size):
                data = sheet.iloc[:,i]
                name = sheet.columns[i]
                if not self.is_should_calculate(option, name,index, data,must_contain):
                    continue
                future_task = executor.submit(put_data,name)
                future_task_list.append(future_task)
            for future in as_completed(future_task_list):
                try:
                    result = future.result()
                    self.exist_uma[result[0]]=result[1]
                    print(f'{result[0]}相性数据读取完成')
                except Exception as e:
                    print(f'任务异常: {e}')
            print('马娘相性数据读取完成')

    def calculate(self):
        print('开始计算马娘历战组合相性')
        combinations_5=list(itertools.combinations(list(self.exist_uma.keys()),5))
        rows=list()
        for combination in combinations_5:
            if self.should_ignore_combination(combination):
                continue
            uma_combination=[self.exist_uma[item] for item in combination]
            for item in ScoreData(uma_combination).getDataList():
                rows.append(item)
        df=pd.DataFrame(rows,columns=['马1','马2','马3','马4','马5','1号相性','2号相性','3号相性','4号相性','5号相性','总相性'])
        df[df['总相性']>500].sort_values(by='总相性' ,kind='quicksort',ascending=False).to_excel(output_file_name,sheet_name='data',index=False,header=True)
        print(f'马娘历战组合相性数据已输出到{output_file_name}')

    def should_ignore_combination(self, combination):
        for item in self.must_contain:
            if item not in combination:
                return True
        return False

    def is_should_calculate(self, option, name,index, data,must_contain):
        if name in must_contain:
            return True
        if 0==data[index['是否拥有']]:
            return False
        if '长' in option and data[index['长']] >'B' :
            return False
        if '中' in option and data[index['中']] >'B' :
            return False
        if '英' in option  and data[index['英']] >'B':
            return False
        if '短' in option and data[index['中']] >'B' :
            return False
        return True

uma_name_set=set(pd.read_excel(file_name,exist_uma_sheet_name).columns[1:])
def is_valid_uma_name(name_set):
    for name in name_set:
        if name not in uma_name_set:
            return False
    return True
option_set = set(['中长','英中','英中长','英短'])
option = input('请输入您的历战选项：中长、英中、英中长、英短\n')

if option not in option_set:
    while True:
        option = input('您输入的历战选项有误，请重新输入您的历战选项：中长、英中、英中长、英短\n')
        if option in option_set:
            break

must_contain=[data.strip() for data in input('请输入您需要指定包含的赛马娘的名称（以逗号分隔多个不同的赛马娘）\n').replace('，',',').split(',') if len(data.strip()) !=0]
if not is_valid_uma_name(must_contain):
        while True:
            must_contain=[data.strip() for data in input('您输入的赛马娘名称有误，请输入您需要指定包含的赛马娘的名称（以逗号分隔多个不同的赛马娘）\n').replace('，',',').split(',')]
            if  is_valid_uma_name(must_contain):
                break
Calculator(option=option,must_contain=set(must_contain)).calculate()
        

