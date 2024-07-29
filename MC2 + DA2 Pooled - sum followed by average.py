#!/usr/bin/env python
# coding: utf-8

# In[6]:


get_ipython().system('pip install pandas')
get_ipython().system('pip install openpyxl')
get_ipython().system('pip install matplotlib')
get_ipython().system('pip install seaborn')
get_ipython().system('pip install scipy')
get_ipython().system('pip install statsmodels')
get_ipython().system('pip install xlsxwriter')


# In[11]:


#sorting the file by brain area

import pandas as pd
file_path = "/Users/sc17237/Desktop/MC2+DA2 Pool/DA2+MC2 copy.xlsx"
df = pd.read_excel(file_path)

with pd.ExcelWriter('/Users/sc17237/Desktop/MC2+DA2 Pool/DA2+MC2 copy.xlsx') as writer:
    for area_name, data in df.groupby('Area Name'):
        data.to_excel(writer, sheet_name=area_name, index=False)


# In[19]:


# Calculating cell density

import pandas as pd

# input
file_path = '/Users/sc17237/Desktop/MC2+DA2 Pool/DA2+MC2 sorted by brain regions.xlsx'
# output
output_path = '/Users/sc17237/Desktop/MC2+DA2 Pool/DA2+MC2 + Cell Density.xlsx'


xls = pd.ExcelFile(file_path)


with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # 遍历所有的tab（sheet）
    for sheet_name in xls.sheet_names:
        # 读取当前sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # use Regular Expression to extract animal batch_animal number as they're from two batches of animals
        df['Mouse_ID'] = df['Image'].str.extract(r'(^.*?_.*?)(?:_.*)')


        
        # 按小鼠ID、Treatment Group和区域名称分组，计算总Cell Count和总Area
        grouped = df.groupby(['Mouse_ID', 'Treatment Group', 'Area Name']).agg({'Cell Count': 'sum', 'Area': 'sum'}).reset_index()
        
        # 计算TRAPed Cell Density（cell per um²）
        grouped['TRAPed Cell Density (cells/um²)'] = grouped['Cell Count'] / grouped['Area']
        
        # 将单位从cell per um²换算成cell per mm²
        grouped['TRAPed Cell Density (cells/mm²)'] = grouped['TRAPed Cell Density (cells/um²)'] * 1e6
        
        # 将结果写入到新的Excel工作表中
        grouped.to_excel(writer, sheet_name=sheet_name, index=False)


# In[9]:


# summary data for each mouse

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 设定图表风格
sns.set(style="whitegrid")

# Excel文件路径
file_path = '/Users/sc17237/Desktop/TRAP2 ket MC2 /MC2_删除BS3+cy3+Test+treatment group+density.xlsx'

# 加载Excel文件中的所有sheet
xls = pd.ExcelFile(file_path)

# 用于存储所有标签页数据的列表
all_data = []

# 遍历所有的tab（sheet）
for sheet_name in xls.sheet_names:
    # 读取当前sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    # 添加Area Name列（因为每个tab代表一个不同的Area Name）
    df['Area Name'] = sheet_name
    # 将当前sheet的数据添加到列表中
    all_data.append(df)

# 合并所有标签页的数据
combined_data = pd.concat(all_data)

# 绘制图表：每只小鼠在不同Area Name的TRAPed Cell Density (cells/mm²)
plt.figure(figsize=(10, 6))  # 可以调整图表大小
chart = sns.barplot(x='Area Name', y='TRAPed Cell Density (cells/mm²)', hue='Mouse_ID', data=combined_data, palette='pastel')
plt.title('TRAPed Cell Density per Area for Each Mouse')
plt.xticks(rotation=45)  # 旋转X轴标签，防止文字重叠
plt.tight_layout()  # 自动调整子图参数, 使之填充整个图像区域
plt.legend(title='Mouse ID', bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)  # 移动图例到图外，防止重叠
plt.show()


# In[22]:


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 设定图表风格
sns.set(style="whitegrid")

# Excel文件路径
file_path = '/Users/sc17237/Desktop/TRAP2 ket DA2 data  重来/删除brainstem slices+treatment group + sorted by area + density copy.xlsx'

# 加载Excel文件中的所有sheet
xls = pd.ExcelFile(file_path)

# 用于存储所有标签页数据的列表
all_data = []

# 遍历所有的tab（sheet）
for sheet_name in xls.sheet_names:
    # 读取当前sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    # 添加Area Name列（因为每个tab代表一个不同的Area Name）
    df['Area Name'] = sheet_name
    # 将当前sheet的数据添加到列表中
    all_data.append(df)

# 合并所有标签页的数据
combined_data = pd.concat(all_data)

# 绘制图表：每只小鼠在不同Area Name的TRAPed Cell Density (cells/mm²)
plt.figure(figsize=(12, 8))  # 增大图表大小
chart = sns.barplot(y='Area Name', x='TRAPed Cell Density (cells/mm²)', hue='Mouse_ID', data=combined_data, palette='pastel', dodge=True)

# 添加平均值标注
for i, area in enumerate(combined_data['Area Name'].unique()):
    # 计算平均值
    mean_val = combined_data[combined_data['Area Name'] == area]['TRAPed Cell Density (cells/mm²)'].mean()
    # 在条形图上方添加平均值标注
    plt.text(y=i, x=mean_val, s=f'{mean_val:.2f}', color='black', va='center')

# 添加分割线
unique_areas = combined_data['Area Name'].nunique()
for i in range(unique_areas - 1):
    plt.axhline(i + 0.5, color='gray', linestyle='--', lw=0.8)

plt.title('TRAPed Cell Density per Area for Each Mouse')
plt.xlabel('TRAPed Cell Density (cells/mm²)')
plt.ylabel('Area Name')
plt.tight_layout()  # 自动调整子图参数, 使之填充整个图像区域
plt.legend(title='Mouse ID', bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)  # 移动图例到图外，防止重叠
plt.show()


# In[20]:


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 设定图表风格
sns.set_style("whitegrid")
plt.rcParams.update({'font.size': 14, 'legend.fontsize': 12})  # 更新字体大小

# 根据你的数据实际情况，更新这个顺序列表
# 假设的从前脑到后脑的区域名称顺序
allen_brain_atlas_order = [
    'OLF', 'CLA', 'ACA', 'PL', 'ILA', 'GU', 'AI', 'MO', 'SS', 'CNU',
    'EP', 'HPF', 'HY', 'LH', 'MH', 'RSP', 'VISC',
    'AUD', 'VIS', 'PTLp', 'TEa', 'PERI', 'ECT', 'DORpm', 'DORsm', 'PA',
    'LA', 'BLA', 'BMA', 'ORB', 'MB'
]

# Excel文件路径
file_path = '/Users/sc17237/Desktop/MC2+DA2 Pool/DA2+MC2 + Cell Density copy.xlsx'

# 加载Excel文件中的所有sheet
xls = pd.ExcelFile(file_path)
all_data = []

# 遍历所有的tab（sheet）
for sheet_name in xls.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df['Area Name'] = sheet_name  # 确保这里正确设置了区域名称
    all_data.append(df)

# 合并所有数据，并设置正确的顺序
combined_data = pd.concat(all_data)
combined_data['Area Name'] = pd.Categorical(combined_data['Area Name'], categories=allen_brain_atlas_order, ordered=True)
combined_data = combined_data.sort_values('Area Name')

# 准备绘图
plt.figure(figsize=(14, 10))
sns.color_palette("vlag", as_cmap=True) # 为每个区域设置不同的颜色
bar_chart = sns.barplot(y='Area Name', x='TRAPed Cell Density (cells/mm²)', data=combined_data,
                        ci=None, palette= "vlag", edgecolor='black', linewidth=1.5)
sns.stripplot(y='Area Name', x='TRAPed Cell Density (cells/mm²)', data=combined_data,
              color='black', size=5, jitter=True, alpha=0.7)  # 使用灰色以增强对比
sns.despine()
plt.title('TRAPed Cell Density per Area MC2')
plt.xlabel('TRAPed Cell Density (cells/mm²)')
plt.ylabel('Area Name')
plt.tight_layout()
plt.show()


# In[15]:


# 创建x轴是animal number的summary data plot
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 设置Seaborn的样式，确保图表美观
sns.set(style="whitegrid")

# 定义文件路径
file_path = '/Users/sc17237/Desktop/TRAP2 ket MC2 /MC2_删除BS3+cy3+Test+treatment group+density.xlsx'

# 读取Excel文件的所有sheets
xls = pd.ExcelFile(file_path)

# 初始化一个空列表来存储所有的DataFrame
dataframes = []

# 遍历所有的sheets
for sheet_name in xls.sheet_names:
    # 读取当前sheet的数据
    data = pd.read_excel(xls, sheet_name=sheet_name)
    # 添加一个新列来存储Area Name
    data['Area Name'] = sheet_name
    # 将当前sheet的DataFrame添加到列表中
    dataframes.append(data)

# 使用pd.concat来合并所有的DataFrame
combined_data = pd.concat(dataframes, ignore_index=True)


# 为了图表清晰，我们可能需要过滤掉一些数据，这里我们保留所有数据
# 你可以根据需要进行调整

# 创建图表
plt.figure(figsize=(15, 10))  # 设置图表大小
# 使用Seaborn的barplot来创建条形图
sns.barplot(x='Mouse_ID', y='TRAPed Cell Density (cells/mm²)', hue='Area Name', data=combined_data, palette='pastel', dodge=True)

# 添加标题和标签
plt.title('TRAPed Cell Density across different Areas MC2', fontsize=15)
plt.xlabel('Mouse ID', fontsize=12)
plt.ylabel('TRAPed Cell Density (cells/mm²)', fontsize=12)

# 优化图表设置以确保清晰度
plt.xticks(rotation=45)  # 旋转X轴标签以避免文字重叠
plt.legend(title='Area Name', bbox_to_anchor=(1.05, 1), loc=2)  # 移动图例以避免覆盖图表

# 显示图表
plt.tight_layout()  # 自动调整子图参数, 使之填充整个图表区域
plt.show()


import numpy as np

# 创建叠加的bar plot
plt.figure(figsize=(15, 10))  # 设置图表大小

# 对于堆叠图，我们需要预处理数据
unique_mice = combined_data['Mouse_ID'].unique()
unique_areas = combined_data['Area Name'].unique()

# 初始化堆叠的基础数据
stacked_values = {mouse: np.zeros(len(unique_areas)) for mouse in unique_mice}
area_indices = {area: idx for idx, area in enumerate(unique_areas)}

# 填充数据
for _, row in combined_data.iterrows():
    stacked_values[row['Mouse_ID']][area_indices[row['Area Name']]] = row['TRAPed Cell Density (cells/mm²)']

# 绘制堆叠条形图
for area_idx, area in enumerate(unique_areas):
    bottom_values = [np.sum(list(stacked_values[mouse_id][:area_idx])) for mouse_id in unique_mice]
    plt.bar(unique_mice, [stacked_values[mouse_id][area_idx] for mouse_id in unique_mice], 
            bottom=bottom_values, label=area)

# 添加图表元素
plt.title('TRAPed Cell Density across Different Areas MC2', fontsize=13)
plt.xlabel('Mouse ID', fontsize=14)
plt.ylabel('TRAPed Cell Density (cells/mm²)', fontsize=14)
plt.xticks(rotation=0)  # 旋转X轴标签以避免文字重叠
plt.legend(title='Area Name', loc='upper right')  # 显示图例

# 显示图表
plt.tight_layout()
plt.show()




# In[22]:


# 不完美但是基本可以用的一个summary whole brain graph

import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
from scipy.stats import f_oneway, shapiro, levene, kruskal
import numpy as np

# 设置Seaborn风格
sns.set(style="whitegrid", palette="deep")

# 加载Excel文件
file_path = '/Users/sc17237/Desktop/MC2+DA2 Pool/DA2+MC2 + Cell Density copy.xlsx' 
xls = pd.ExcelFile(file_path)

# 定义排序顺序
allen_brain_atlas_order = [
    'OLF', 'CLA', 'ACA', 'PL', 'ILA', 'GU', 'AI', 'MO', 'SS', 'CNU',
    'EP', 'HPF', 'HY', 'LH', 'MH', 'RSP', 'VISC',
    'AUD', 'VIS', 'PTLp', 'TEa', 'PERI', 'ECT', 'DORpm', 'DORsm', 'PA',
    'LA', 'BLA', 'BMA', 'ORB', 'MB'
]

# 按照给定的顺序排序工作表名
sorted_sheet_names = sorted(xls.sheet_names, key=lambda x: allen_brain_atlas_order.index(x))

# 准备绘图
fig, axes = plt.subplots(len(sorted_sheet_names), 2, figsize=(30, 50), gridspec_kw={'width_ratios': [4, 1]})

# 初始化存储P值的列表
p_values = []

# 遍历所有的区域，绘制每个区域的数据
for i, area in enumerate(sorted_sheet_names):
    area_data = pd.read_excel(xls, sheet_name=area)
    # 重新排序Treatment Group列以确保条形图的顺序
    # 首先，定义新顺序
    treatment_order = ['Vehicle', '1 mg', '5 mg']
    # 使用Categorical数据类型来确保顺序
    area_data['Treatment Group'] = pd.Categorical(area_data['Treatment Group'], categories=treatment_order, ordered=True)
    # 按新顺序对数据进行排序
    area_data.sort_values('Treatment Group', inplace=True)
    # 计算每个治疗组的平均值和标准差
    group_stats = area_data.groupby('Treatment Group')['TRAPed Cell Density (cells/mm²)'].agg(['mean', 'std', 'count'])
    group_stats['sem'] = group_stats['std'] / np.sqrt(group_stats['count'])  # 标准误
    
    # 绘制条形图
    ax = axes[i, 0]
    sns.color_palette("Set2")
    sns.barplot(data=group_stats.reset_index(), x='mean', y='Treatment Group', ax=ax, palette="Set2", ci=None)
    ax.errorbar(x=group_stats['mean'], y=group_stats.index, xerr=group_stats['sem'], fmt='none', c='black', capsize=5)
    
    # 添加每个数据点
    sns.stripplot(data=area_data, x='TRAPed Cell Density (cells/mm²)', y='Treatment Group', ax=ax, color='black', alpha=0.5, jitter=True)
    
    # 统计检验
    groups = area_data.groupby('Treatment Group')['TRAPed Cell Density (cells/mm²)'].apply(list).to_dict()
    vehicle = groups.get('Vehicle', [])
    one_mg = groups.get('1 mg', [])
    five_mg = groups.get('5 mg', [])
    if all(shapiro(data)[1] > 0.05 for data in groups.values()) and levene(vehicle, one_mg, five_mg)[1] > 0.05:
        _, p_value = f_oneway(vehicle, one_mg, five_mg)
    else:
        _, p_value = kruskal(vehicle, one_mg, five_mg)
    p_values.append(p_value)
    
    # 标注区域名称
    ax.set_ylabel(area)
    

        # 清除x轴标签
    if i != len(sorted_sheet_names) - 1:
        ax.set_xticks([])  # 移除非最后一行的x轴刻度
        ax.set_xlabel('')  # 移除非最后一行的x轴标签
    else:
        # 仅在最后一行设置x轴标签
        ax.set_xlabel('Mean TRAPed Cell Density MC2')
        
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

# 设置所有条形图的相同x轴范围
max_mean = max(group_stats['mean'] + group_stats['sem'])
for ax in axes[:, 0]:
    ax.set_xlim(0, max_mean + max_mean * 1.2)

# 在最右侧添加一张显示所有P值的条形图
for i, p_value in enumerate(p_values):
    ax = axes[i, 1]
    ax.barh(y=[0], width=[p_value], color='grey')
    ax.set_xlim(0, 1)
    ax.set_ylim(-0.5, 0.5)
    ax.set_yticks([])
    if i != len(sorted_sheet_names) - 1:
        ax.set_xticks([]) 
    
    else:
        # 仅在最后一行设置x轴标签
        ax.set_xlabel('P value')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

# 调整布局
#plt.tight_layout()
plt.show()


# In[21]:


# key ROIs graph

import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
from scipy.stats import f_oneway, shapiro, levene, kruskal
import numpy as np

# 设置Seaborn风格
sns.set(style="whitegrid", palette="deep")

# 加载Excel文件
file_path = '/Users/sc17237/Desktop/MC2+DA2 Pool/DA2+MC2 + Cell Density copy.xlsx'  # 请替换为实际路径
xls = pd.ExcelFile(file_path)

# 定义排序顺序
allen_brain_atlas_order = [
    'ACA', 'PL', 'ILA', 'LH', 'BLA'
]

# 按照给定的顺序排序工作表名，仅包括存在于allen_brain_atlas_order中的工作表
sorted_sheet_names = sorted([sheet for sheet in xls.sheet_names if sheet in allen_brain_atlas_order],
                            key=lambda x: allen_brain_atlas_order.index(x))

# 准备绘图
fig, axes = plt.subplots(len(sorted_sheet_names), 2, figsize=(10, 20), gridspec_kw={'width_ratios': [4, 1]})

# 初始化存储P值的列表
p_values = []

# 遍历所有的区域，绘制每个区域的数据
for i, area in enumerate(sorted_sheet_names):
    area_data = pd.read_excel(xls, sheet_name=area)
     # 重新排序Treatment Group列以确保条形图的顺序
    # 首先，定义新顺序
    treatment_order = ['Vehicle', '1 mg', '5 mg']
    # 使用Categorical数据类型来确保顺序
    area_data['Treatment Group'] = pd.Categorical(area_data['Treatment Group'], categories=treatment_order, ordered=True)
    # 按新顺序对数据进行排序
    area_data.sort_values('Treatment Group', inplace=True)
    
    # 计算每个治疗组的平均值和标准差
    group_stats = area_data.groupby('Treatment Group')['TRAPed Cell Density (cells/mm²)'].agg(['mean', 'std', 'count'])
    group_stats['sem'] = group_stats['std'] / np.sqrt(group_stats['count'])  # 标准误
    
    # 绘制条形图
    ax = axes[i, 0]
    sns.color_palette("Set2")
    sns.barplot(data=group_stats.reset_index(), x='mean', y='Treatment Group', ax=ax, palette="Set2", ci=None)
    ax.errorbar(x=group_stats['mean'], y=group_stats.index, xerr=group_stats['sem'], fmt='none', c='black', capsize=5)
    
    # 添加每个数据点
    sns.stripplot(data=area_data, x='TRAPed Cell Density (cells/mm²)', y='Treatment Group', ax=ax, color='black', alpha=0.5, jitter=True)
    
    # 统计检验
    groups = area_data.groupby('Treatment Group')['TRAPed Cell Density (cells/mm²)'].apply(list).to_dict()
    vehicle = groups.get('Vehicle', [])
    one_mg = groups.get('1 mg', [])
    five_mg = groups.get('5 mg', [])
    if all(shapiro(data)[1] > 0.05 for data in groups.values()) and levene(vehicle, one_mg, five_mg)[1] > 0.05:
        _, p_value = f_oneway(vehicle, one_mg, five_mg)
    else:
        _, p_value = kruskal(vehicle, one_mg, five_mg)
    p_values.append(p_value)
    
    # 标注区域名称
    ax.set_ylabel(area)
    

        # 清除x轴标签
    if i != len(sorted_sheet_names) - 1:
        ax.set_xticks([])  # 移除非最后一行的x轴刻度
        ax.set_xlabel('')  # 移除非最后一行的x轴标签
    else:
        # 仅在最后一行设置x轴标签
        ax.set_xlabel('Mean TRAPed Cell Density MC2')
        
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

# 设置所有条形图的相同x轴范围
max_mean = max(group_stats['mean'] + group_stats['sem'])
for ax in axes[:, 0]:
    ax.set_xlim(0, max_mean + max_mean * 1.7)

# 在最右侧添加一张显示所有P值的条形图
for i, p_value in enumerate(p_values):
    ax = axes[i, 1]
    ax.barh(y=[0], width=[p_value], color='grey')
    ax.set_xlim(0, 1)
    ax.set_ylim(-0.5, 0.5)
    ax.set_yticks([])
    # 在条形图上方显示P值
    ax.text(p_value + 0.05, 0, f'P={p_value:.3f}', va='center')
    if i != len(sorted_sheet_names) - 1:
        ax.set_xticks([]) 
    
    else:
        # 仅在最后一行设置x轴标签
        ax.set_xlabel('P value')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(True)
    ax.spines['bottom'].set_visible(True)

# 调整布局
#plt.tight_layout()
plt.show()


# In[2]:


import pandas as pd
import statsmodels.api as sm
from statsmodels.formula.api import ols

# input excel
file_path = '/Users/sc17237/Desktop/TRAP Pool 2Way ANOVA/DA2+MC2 + Cell Density + 2way ANOVA.xlsx'

# read all sheets for each ROIs
all_sheets = pd.read_excel(file_path, sheet_name=None)

ancova_results_dict = {}


for sheet_name, sheet_data in all_sheets.items():
    formula_sheet = 'Q("TRAPed Cell Density (cells/mm²)") ~ C(Q("Treatment Group")) + C(Sex) + C(Batch)'
    model_sheet = ols(formula_sheet, data=sheet_data).fit()
    
   
    ancova_results_sheet = sm.stats.anova_lm(model_sheet, typ=2)
    ancova_results_dict[sheet_name] = ancova_results_sheet


output_path = '/Users/sc17237/Desktop/TRAP Pool 2Way ANOVA/DA2+MC2 + Cell Density + ANCOVA.xlsx'
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, results in ancova_results_dict.items():
        results.to_excel(writer, sheet_name=sheet_name)


print(f'Results saved to {output_path}')


# In[3]:


import pandas as pd
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd

# 路径可能需要根据你的文件位置进行调整
file_path = '/Users/sc17237/Desktop/TRAP Pool 2Way ANOVA/DA2+MC2 + Cell Density + 2way ANOVA.xlsx'

# 读取Excel文件的所有工作表
all_sheets = pd.read_excel(file_path, sheet_name=None)

# 字典用于存储每个工作表的ANCOVA结果
ancova_results_dict = {}
# 字典用于存储每个工作表的Tukey HSD结果
tukey_results_dict = {}

# 处理每个工作表
for sheet_name, sheet_data in all_sheets.items():
    # 准备并拟合当前工作表的模型
    formula_sheet = 'Q("TRAPed Cell Density (cells/mm²)") ~ C(Q("Treatment Group")) + C(Sex) + C(Batch)'
    model_sheet = ols(formula_sheet, data=sheet_data).fit()
    
    # 执行ANCOVA并存储结果
    ancova_results_sheet = sm.stats.anova_lm(model_sheet, typ=2)
    ancova_results_dict[sheet_name] = ancova_results_sheet

    # 执行Tukey HSD测试
    tukey = pairwise_tukeyhsd(endog=sheet_data['TRAPed Cell Density (cells/mm²)'], 
                              groups=sheet_data['Treatment Group'], 
                              alpha=0.05)
    tukey_results_dict[sheet_name] = tukey.summary()

# 创建新的Excel写入器并保存每个结果到单独的工作表
output_path = '/Users/sc17237/Desktop/TRAP Pool 2Way ANOVA/DA2+MC2 + Cell Density + ANCOVA + post hoc.xlsx'
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, results in ancova_results_dict.items():
        results.to_excel(writer, sheet_name=sheet_name + "_ANCOVA")
    for sheet_name, tukey_results in tukey_results_dict.items():
        tukey_results_df = pd.DataFrame(data=tukey_results.data[1:], columns=tukey_results.data[0])
        tukey_results_df.to_excel(writer, sheet_name=sheet_name + "_TukeyHSD")

# 输出文件路径，根据需要调整
print(f'Results saved to {output_path}')


# In[ ]:




